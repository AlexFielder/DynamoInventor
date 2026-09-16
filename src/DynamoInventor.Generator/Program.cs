using System.Reflection;
using System.Text;
using System.Text.Json;

namespace DynamoInventor.Generator;

/// <summary>
/// Emits InventorLibrary.API.Inv* wrapper classes from the installed Inventor interop.
///
/// Shape of the output (kept compatible with the original IronPython generator so hand-written
/// nodes keep working): a partial class per interop interface with
///   internal Inventor.X InternalX                  the wrapped COM object
///   public   Inventor.X XInstance                  same, public (excluded from the Dynamo VM)
///   internal T InternalProp / public T Prop        properties, wrapper-typed where the type is wrapped
///   private  T InternalMethod() / public T Method  methods
///   public static InvX ByInvX(InvX) / ByInvX(Inventor.X)
/// and an enum per interop enum. Members whose types are not in the wrapped set are emitted as
/// comments, never as raw interop types, so the generated surface stays compilable and does not
/// leak Inventor.* type names into DesignScript.
/// </summary>
internal static class Program
{
    private sealed record Config(string Namespace, string Prefix, string InteropNamespace, string[] Types);

    private static int Main(string[] args)
    {
        string? configPath = Arg(args, "--config");
        if (configPath == null)
        {
            Console.Error.WriteLine("usage: DynamoInventor.Generator --config <wrappers.json> [--interop <Autodesk.Inventor.Interop.dll>] [--out <dir>] [--dry-run]");
            return 2;
        }

        var config = JsonSerializer.Deserialize<Config>(File.ReadAllText(configPath),
            new JsonSerializerOptions { PropertyNameCaseInsensitive = true })
            ?? throw new InvalidOperationException("Could not read " + configPath);

        var interopPath = Arg(args, "--interop") ?? FindInterop();
        var outDir = Arg(args, "--out") ?? Path.GetDirectoryName(Path.GetFullPath(configPath))!;
        var dryRun = args.Contains("--dry-run");

        // Interop signatures reference stdole (IPictureDisp etc.), which ships in Inventor's Bin folder
        // one level above Public Assemblies; resolve it and any other neighbour from there.
        var probeFolders = new[] { Path.GetDirectoryName(interopPath)!, Path.GetFullPath(Path.Combine(Path.GetDirectoryName(interopPath)!, "..")) };
        System.Runtime.Loader.AssemblyLoadContext.Default.Resolving += (ctx, name) =>
        {
            foreach (var folder in probeFolders)
            {
                var candidate = Path.Combine(folder, name.Name + ".dll");
                if (File.Exists(candidate)) return ctx.LoadFromAssemblyPath(candidate);
            }
            return null;
        };

        var interop = Assembly.LoadFrom(interopPath);
        Console.WriteLine($"Interop {interop.GetName().Version} from {interopPath}");
        Console.WriteLine($"Output  {outDir}{(dryRun ? " (dry run)" : "")}");

        var wrapped = new HashSet<string>(config.Types, StringComparer.Ordinal);
        var mapper = new TypeMapper(config.InteropNamespace, config.Prefix, wrapped);
        var stats = new Stats();
        var missing = new List<string>();

        foreach (var name in config.Types)
        {
            var type = interop.GetType(config.InteropNamespace + "." + name);
            if (type == null)
            {
                missing.Add(name);
                continue;
            }

            var text = type.IsEnum ? EmitEnum(type, config) : EmitClass(type, config, mapper, stats);
            var file = Path.Combine(outDir, config.Prefix + name + ".cs");
            if (!dryRun)
            {
                File.WriteAllText(file, text, new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));
            }
            stats.Files++;
        }

        Console.WriteLine($"{stats.Files} files, {stats.Members} members emitted, {stats.Skipped} members skipped (types outside the wrapped set).");
        if (missing.Count > 0)
        {
            Console.Error.WriteLine("Not found in interop: " + string.Join(", ", missing));
            return 1;
        }
        return 0;
    }

    // ---------------------------------------------------------------- enums

    private static string EmitEnum(Type type, Config config)
    {
        var sb = new StringBuilder();
        Header(sb, type, config);
        sb.AppendLine("    [IsVisibleInDynamoLibrary(false)]");
        sb.AppendLine($"    public enum {config.Prefix}{type.Name}");
        sb.AppendLine("    {");
        var names = Enum.GetNames(type).OrderBy(n => n, StringComparer.Ordinal).ToArray();
        for (var i = 0; i < names.Length; i++)
        {
            sb.AppendLine($"        {names[i]} = {config.InteropNamespace}.{type.Name}.{names[i]}{(i < names.Length - 1 ? "," : "")}");
        }
        sb.AppendLine("    }");
        sb.AppendLine("}");
        return sb.ToString();
    }

    // -------------------------------------------------------------- classes

    private static string EmitClass(Type type, Config config, TypeMapper mapper, Stats stats)
    {
        var raw = config.InteropNamespace + "." + type.Name;      // Inventor.WorkPoint
        var wrapper = config.Prefix + type.Name;                    // InvWorkPoint
        var instance = type.Name + "Instance";                      // WorkPointInstance
        var internalField = "Internal" + type.Name;                 // InternalWorkPoint

        // Some interop interfaces declare a member twice (e.g. a read-only and a read-write Name from
        // different dispinterface versions); keep one per name, preferring the writable property.
        var props = type.GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .Where(p => !p.Name.StartsWith('_'))
            .GroupBy(p => p.Name, StringComparer.Ordinal)
            .Select(g => g.FirstOrDefault(p => p.CanWrite) ?? g.First())
            .OrderBy(p => p.Name, StringComparer.Ordinal)
            .ToList();
        var methods = type.GetMethods(BindingFlags.Public | BindingFlags.Instance)
            .Where(m => !m.IsSpecialName && !m.Name.StartsWith('_'))
            .GroupBy(m => m.Name + "(" + string.Join(",", m.GetParameters().Select(a => a.ParameterType.FullName)) + ")", StringComparer.Ordinal)
            .Select(g => g.First())
            .OrderBy(m => m.Name, StringComparer.Ordinal)
            .ToList();

        // The interop occasionally owns a name our "Internal" twin would take (Material.InternalName);
        // suffix the twin in that case so the two never collide.
        var memberNames = new HashSet<string>(props.Select(p => p.Name).Concat(methods.Select(m => m.Name)), StringComparer.Ordinal);
        string InternalName(string member) => memberNames.Contains("Internal" + member) ? "Internal" + member + "Raw" : "Internal" + member;

        var sb = new StringBuilder();
        Header(sb, type, config);
        sb.AppendLine("    [IsVisibleInDynamoLibrary(false)]");
        sb.AppendLine($"    public partial class {wrapper}");
        sb.AppendLine("    {");

        // --- internal side -------------------------------------------------
        sb.AppendLine("        #region Internal properties");
        sb.AppendLine($"        internal {raw} {internalField} {{ get; set; }}");
        sb.AppendLine();
        var internalProps = new StringBuilder();
        var publicProps = new StringBuilder();
        foreach (var p in props)
        {
            EmitProperty(p, instance, InternalName(p.Name), mapper, stats, internalProps, publicProps);
        }
        sb.Append(internalProps);
        sb.AppendLine("        #endregion");
        sb.AppendLine();

        sb.AppendLine("        #region Private constructors");
        sb.AppendLine($"        private {wrapper}({wrapper} {Camel(wrapper)})");
        sb.AppendLine("        {");
        sb.AppendLine($"            {internalField} = {Camel(wrapper)}.{internalField};");
        sb.AppendLine("        }");
        sb.AppendLine();
        sb.AppendLine($"        private {wrapper}({raw} {Camel(wrapper)})");
        sb.AppendLine("        {");
        sb.AppendLine($"            {internalField} = {Camel(wrapper)};");
        sb.AppendLine("        }");
        sb.AppendLine("        #endregion");
        sb.AppendLine();

        var privateMethods = new StringBuilder();
        var publicMethods = new StringBuilder();
        foreach (var m in methods)
        {
            EmitMethod(m, instance, InternalName(m.Name), mapper, stats, privateMethods, publicMethods);
        }
        sb.AppendLine("        #region Private methods");
        sb.Append(privateMethods);
        sb.AppendLine("        #endregion");
        sb.AppendLine();

        // --- public side ---------------------------------------------------
        sb.AppendLine("        #region Public properties");
        sb.AppendLine("        /// <summary>The wrapped Inventor COM object. Not imported into the Dynamo VM.</summary>");
        sb.AppendLine("        [SupressImportIntoVM]");
        sb.AppendLine($"        public {raw} {instance}");
        sb.AppendLine("        {");
        sb.AppendLine($"            get {{ return {internalField}; }}");
        sb.AppendLine($"            set {{ {internalField} = value; }}");
        sb.AppendLine("        }");
        sb.AppendLine();
        sb.Append(publicProps);
        sb.AppendLine("        #endregion");
        sb.AppendLine();

        sb.AppendLine("        #region Public static constructors");
        sb.AppendLine($"        public static {wrapper} By{wrapper}({wrapper} {Camel(wrapper)})");
        sb.AppendLine("        {");
        sb.AppendLine($"            return {Camel(wrapper)} == null ? null : new {wrapper}({Camel(wrapper)});");
        sb.AppendLine("        }");
        sb.AppendLine();
        sb.AppendLine("        [SupressImportIntoVM]");
        sb.AppendLine($"        public static {wrapper} By{wrapper}({raw} {Camel(wrapper)})");
        sb.AppendLine("        {");
        sb.AppendLine($"            return {Camel(wrapper)} == null ? null : new {wrapper}({Camel(wrapper)});");
        sb.AppendLine("        }");
        sb.AppendLine("        #endregion");
        sb.AppendLine();

        sb.AppendLine("        #region Public methods");
        sb.Append(publicMethods);
        sb.AppendLine("        #endregion");
        sb.AppendLine("    }");
        sb.AppendLine("}");
        return sb.ToString();
    }

    private static void EmitProperty(PropertyInfo p, string instance, string internalName, TypeMapper mapper, Stats stats,
        StringBuilder internalOut, StringBuilder publicOut)
    {
        var indexParams = p.GetIndexParameters();
        if (!mapper.TryMap(p.PropertyType, out var mapped))
        {
            Skip(internalOut, $"{p.Name}: {Describe(p.PropertyType)} is not wrapped", stats);
            return;
        }

        if (indexParams.Length > 0)
        {
            // Parameterised properties become methods named after the property. C# can only spell the
            // default indexer ("Item") with [] syntax; any other parameterised property is "not supported
            // by the language" and must go through its get_ accessor.
            if (!TryMapParameters(indexParams, mapper, out var sig, out var call, out var why)
                || !TryMapParameters(indexParams, mapper, out _, out var forward, out _, forwardOnly: true))
            {
                Skip(internalOut, $"indexer {p.Name}: {why}", stats);
                return;
            }
            // Only the type's DefaultMember indexer can be written with []; NameValueMap.Item, for one, is not.
            var isDefaultIndexer = p.DeclaringType?.GetCustomAttribute<System.Reflection.DefaultMemberAttribute>()?.MemberName == p.Name;
            var rawAccess = isDefaultIndexer ? $"{instance}[{call}]" : $"{instance}.get_{p.Name}({call})";
            internalOut.AppendLine($"        internal {mapped.CSharp} {internalName}({sig})");
            internalOut.AppendLine("        {");
            internalOut.AppendLine($"            return {mapped.FromRaw(rawAccess)};");
            internalOut.AppendLine("        }");
            internalOut.AppendLine();
            publicOut.AppendLine($"        public {mapped.CSharp} {p.Name}({sig})");
            publicOut.AppendLine("        {");
            publicOut.AppendLine($"            return {internalName}({forward});");
            publicOut.AppendLine("        }");
            publicOut.AppendLine();
            stats.Members++;
            return;
        }

        var setter = p.GetSetMethod();
        var writable = p.CanWrite && setter != null;

        // A COM property whose setter takes a different type from the getter's return (Parameter.Units:
        // get string / set object) cannot be a C# property. Expose the getter as a read-only property
        // and the setter as a Set<Name>(value) method through the accessor methods.
        if (writable && setter!.GetParameters().Length == 1 && setter.GetParameters()[0].ParameterType != p.PropertyType)
        {
            if (!mapper.TryMap(setter.GetParameters()[0].ParameterType, out var setMapped))
            {
                Skip(internalOut, $"{p.Name} setter: {Describe(setter.GetParameters()[0].ParameterType)} is not wrapped", stats);
                return;
            }
            internalOut.AppendLine($"        internal {mapped.CSharp} {internalName}");
            internalOut.AppendLine("        {");
            internalOut.AppendLine($"            get {{ return {mapped.FromRaw($"{instance}.get_{p.Name}()")}; }}");
            internalOut.AppendLine("        }");
            internalOut.AppendLine();
            internalOut.AppendLine($"        internal void {internalName}Set({setMapped.CSharp} value)");
            internalOut.AppendLine("        {");
            internalOut.AppendLine($"            {instance}.set_{p.Name}({setMapped.ToRaw("value")});");
            internalOut.AppendLine("        }");
            internalOut.AppendLine();
            publicOut.AppendLine($"        public {mapped.CSharp} {p.Name}");
            publicOut.AppendLine("        {");
            publicOut.AppendLine($"            get {{ return {internalName}; }}");
            publicOut.AppendLine("        }");
            publicOut.AppendLine();
            publicOut.AppendLine($"        public void Set{p.Name}({setMapped.CSharp} value)");
            publicOut.AppendLine("        {");
            publicOut.AppendLine($"            {internalName}Set(value);");
            publicOut.AppendLine("        }");
            publicOut.AppendLine();
            stats.Members += 2;
            return;
        }
        internalOut.AppendLine($"        internal {mapped.CSharp} {internalName}");
        internalOut.AppendLine("        {");
        internalOut.AppendLine($"            get {{ return {mapped.FromRaw($"{instance}.{p.Name}")}; }}");
        if (writable)
        {
            internalOut.AppendLine($"            set {{ {instance}.{p.Name} = {mapped.ToRaw("value")}; }}");
        }
        internalOut.AppendLine("        }");
        internalOut.AppendLine();

        publicOut.AppendLine($"        public {mapped.CSharp} {p.Name}");
        publicOut.AppendLine("        {");
        publicOut.AppendLine($"            get {{ return {internalName}; }}");
        if (writable)
        {
            publicOut.AppendLine($"            set {{ {internalName} = value; }}");
        }
        publicOut.AppendLine("        }");
        publicOut.AppendLine();
        stats.Members++;
    }

    private static void EmitMethod(MethodInfo m, string instance, string internalName, TypeMapper mapper, Stats stats,
        StringBuilder privateOut, StringBuilder publicOut)
    {
        if (!mapper.TryMap(m.ReturnType, out var ret))
        {
            Skip(privateOut, $"{m.Name}(): returns {Describe(m.ReturnType)}, not wrapped", stats);
            return;
        }
        if (!TryMapParameters(m.GetParameters(), mapper, out var sig, out var call, out var why))
        {
            Skip(privateOut, $"{m.Name}(): {why}", stats);
            return;
        }

        var isVoid = m.ReturnType == typeof(void);
        var forward = TryMapParameters(m.GetParameters(), mapper, out _, out var forwardCall, out _, forwardOnly: true) ? forwardCall : call;

        privateOut.AppendLine($"        private {ret.CSharp} {internalName}({sig})");
        privateOut.AppendLine("        {");
        privateOut.AppendLine(isVoid
            ? $"            {instance}.{m.Name}({call});"
            : $"            return {ret.FromRaw($"{instance}.{m.Name}({call})")};");
        privateOut.AppendLine("        }");
        privateOut.AppendLine();

        publicOut.AppendLine($"        public {ret.CSharp} {m.Name}({sig})");
        publicOut.AppendLine("        {");
        publicOut.AppendLine(isVoid
            ? $"            {internalName}({forward});"
            : $"            return {internalName}({forward});");
        publicOut.AppendLine("        }");
        publicOut.AppendLine();
        stats.Members++;
    }

    /// <summary>
    /// Builds the C# parameter list and the argument list used to call the COM member.
    /// With forwardOnly the argument list forwards our own parameters unchanged (public -> Internal).
    /// </summary>
    private static bool TryMapParameters(ParameterInfo[] parameters, TypeMapper mapper,
        out string signature, out string arguments, out string why, bool forwardOnly = false)
    {
        var sig = new List<string>();
        var call = new List<string>();
        foreach (var p in parameters)
        {
            var t = p.ParameterType;
            var byRef = t.IsByRef;
            var elem = byRef ? t.GetElementType()! : t;
            if (!mapper.TryMap(elem, out var mapped))
            {
                signature = arguments = string.Empty;
                why = $"parameter {p.Name} is {Describe(elem)}, not wrapped";
                return false;
            }
            var mod = byRef ? (p.IsOut && !p.IsIn ? "out " : "ref ") : string.Empty;
            if (byRef && mapped.Kind != Kind.Builtin)
            {
                signature = arguments = string.Empty;
                why = $"parameter {p.Name} is a ref/out {Describe(elem)}";
                return false;
            }
            var name = SafeName(Camel(p.Name ?? "arg"));
            sig.Add($"{mod}{mapped.CSharp} {name}");
            call.Add(byRef || forwardOnly ? $"{mod}{name}" : mapped.ToRaw(name));
        }
        signature = string.Join(", ", sig);
        arguments = string.Join(", ", call);
        why = string.Empty;
        return true;
    }

    // -------------------------------------------------------------- helpers

    private static void Header(StringBuilder sb, Type type, Config config)
    {
        sb.AppendLine("// <auto-generated>");
        sb.AppendLine($"//     Generated by DynamoInventor.Generator from {type.Assembly.GetName().Name} {type.Assembly.GetName().Version}");
        sb.AppendLine("//     for " + config.InteropNamespace + "." + type.Name + ". Do not edit; put additions in " + config.Prefix + type.Name + ".Custom.cs.");
        sb.AppendLine("// </auto-generated>");
        sb.AppendLine("using System;");
        sb.AppendLine("using System.Collections.Generic;");
        sb.AppendLine("using Autodesk.DesignScript.Runtime;");
        sb.AppendLine();
        sb.AppendLine($"namespace {config.Namespace}");
        sb.AppendLine("{");
    }

    private static void Skip(StringBuilder sb, string reason, Stats stats)
    {
        sb.AppendLine($"        // skipped: {reason}");
        stats.Skipped++;
    }

    private static string Describe(Type t) => t.IsByRef ? Describe(t.GetElementType()!) + "&" : (t.FullName ?? t.Name);

    private static string Camel(string s) => string.IsNullOrEmpty(s) ? s : char.ToLowerInvariant(s[0]) + s.Substring(1);

    private static readonly HashSet<string> Keywords = new(StringComparer.Ordinal)
    {
        "abstract","as","base","bool","break","byte","case","catch","char","checked","class","const","continue","decimal",
        "default","delegate","do","double","else","enum","event","explicit","extern","false","finally","fixed","float","for",
        "foreach","goto","if","implicit","in","int","interface","internal","is","lock","long","namespace","new","null","object",
        "operator","out","override","params","private","protected","public","readonly","ref","return","sbyte","sealed","short",
        "sizeof","stackalloc","static","string","struct","switch","this","throw","true","try","typeof","uint","ulong","unchecked",
        "unsafe","ushort","using","virtual","void","volatile","while"
    };

    private static string SafeName(string name) => Keywords.Contains(name) ? "@" + name : name;

    private static string? Arg(string[] args, string name)
    {
        for (var i = 0; i < args.Length - 1; i++)
        {
            if (string.Equals(args[i], name, StringComparison.OrdinalIgnoreCase)) return args[i + 1];
        }
        return null;
    }

    private static string FindInterop()
    {
        var root = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Autodesk");
        var candidates = Directory.Exists(root)
            ? Directory.GetDirectories(root, "Inventor 2*")
                .OrderByDescending(d => d, StringComparer.OrdinalIgnoreCase)
                .Select(d => Path.Combine(d, "Bin", "Public Assemblies", "Autodesk.Inventor.Interop.dll"))
                .Where(File.Exists)
                .ToList()
            : new List<string>();
        return candidates.FirstOrDefault()
            ?? throw new FileNotFoundException("Autodesk.Inventor.Interop.dll not found under Program Files\\Autodesk\\Inventor 2*; pass --interop.");
    }

    private sealed class Stats
    {
        public int Files;
        public int Members;
        public int Skipped;
    }
}

internal enum Kind { Builtin, Wrapper, Enum }

/// <summary>A C# type name plus the expressions that convert to and from the raw COM value.</summary>
internal sealed record Mapped(string CSharp, Kind Kind, string RawName)
{
    public string FromRaw(string expr) => Kind switch
    {
        Kind.Wrapper => $"{CSharp}.By{CSharp}({expr})",
        Kind.Enum => $"({CSharp}){expr}",
        _ => expr
    };

    public string ToRaw(string expr) => Kind switch
    {
        Kind.Wrapper => $"{expr}?.{RawName.Substring(RawName.LastIndexOf('.') + 1)}Instance",
        Kind.Enum => $"({RawName}){expr}",
        _ => expr
    };
}

internal sealed class TypeMapper
{
    private static readonly Dictionary<Type, string> Builtins = new()
    {
        [typeof(bool)] = "bool", [typeof(byte)] = "byte", [typeof(char)] = "char", [typeof(decimal)] = "decimal",
        [typeof(double)] = "double", [typeof(short)] = "short", [typeof(int)] = "int", [typeof(long)] = "long",
        [typeof(sbyte)] = "sbyte", [typeof(float)] = "float", [typeof(string)] = "string", [typeof(ushort)] = "ushort",
        [typeof(uint)] = "uint", [typeof(ulong)] = "ulong", [typeof(void)] = "void", [typeof(object)] = "object",
        [typeof(DateTime)] = "System.DateTime", [typeof(IntPtr)] = "System.IntPtr"
    };

    private readonly string interopNamespace;
    private readonly string prefix;
    private readonly HashSet<string> wrapped;

    public TypeMapper(string interopNamespace, string prefix, HashSet<string> wrapped)
    {
        this.interopNamespace = interopNamespace;
        this.prefix = prefix;
        this.wrapped = wrapped;
    }

    public bool TryMap(Type t, out Mapped mapped)
    {
        if (t.IsByRef) t = t.GetElementType()!;

        if (Builtins.TryGetValue(t, out var alias))
        {
            mapped = new Mapped(alias, Kind.Builtin, alias);
            return true;
        }
        if (t.IsArray && t.GetElementType() is { } elem && Builtins.TryGetValue(elem, out var elemAlias) && t.GetArrayRank() == 1)
        {
            mapped = new Mapped(elemAlias + "[]", Kind.Builtin, elemAlias + "[]");
            return true;
        }
        if (t.Namespace == interopNamespace && wrapped.Contains(t.Name))
        {
            var raw = interopNamespace + "." + t.Name;
            mapped = new Mapped(prefix + t.Name, t.IsEnum ? Kind.Enum : Kind.Wrapper, raw);
            return true;
        }
        mapped = null!;
        return false;
    }
}
