using System;
using Inventor;

namespace InventorServices.Utilities
{
    /// <summary>
    /// The Inventor interop's base ComponentDefinition interface does not declare WorkPoints,
    /// WorkPlanes or Parameters; PartComponentDefinition and AssemblyComponentDefinition each
    /// redeclare them. These helpers switch on the concrete interface so nodes can be written once
    /// for both document types.
    /// </summary>
    public static class ComponentDefinitionExtensions
    {
        public static WorkPoints GetWorkPoints(this ComponentDefinition definition)
        {
            switch (definition)
            {
                case PartComponentDefinition part: return part.WorkPoints;
                case AssemblyComponentDefinition assembly: return assembly.WorkPoints;
                default: throw Unsupported(definition);
            }
        }

        public static WorkPlanes GetWorkPlanes(this ComponentDefinition definition)
        {
            switch (definition)
            {
                case PartComponentDefinition part: return part.WorkPlanes;
                case AssemblyComponentDefinition assembly: return assembly.WorkPlanes;
                default: throw Unsupported(definition);
            }
        }

        public static WorkAxes GetWorkAxes(this ComponentDefinition definition)
        {
            switch (definition)
            {
                case PartComponentDefinition part: return part.WorkAxes;
                case AssemblyComponentDefinition assembly: return assembly.WorkAxes;
                default: throw Unsupported(definition);
            }
        }

        public static Parameters GetParameters(this ComponentDefinition definition)
        {
            switch (definition)
            {
                case PartComponentDefinition part: return part.Parameters;
                case AssemblyComponentDefinition assembly: return assembly.Parameters;
                default: throw Unsupported(definition);
            }
        }

        /// <summary>The owning document (PartDocument or AssemblyDocument) as the common Document interface.</summary>
        public static Document GetDocument(this ComponentDefinition definition)
        {
            return (Document)definition.Document;
        }

        private static InvalidOperationException Unsupported(ComponentDefinition definition)
        {
            return new InvalidOperationException(
                "Unsupported component definition type " + (definition == null ? "null" : definition.Type.ToString()) +
                "; only parts and assemblies are supported.");
        }
    }
}
