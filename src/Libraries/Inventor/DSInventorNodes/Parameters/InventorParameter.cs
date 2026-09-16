using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.DesignScript.Runtime;
using Inventor;
using InventorServices.Persistence;
using InventorServices.Utilities;

namespace InventorLibrary.Parameters
{
    /// <summary>
    /// Reads and drives the parameters of the active part or assembly, the same ones iLogic and the
    /// Parameters dialog show. Values are in the parameter's own units; expressions are Inventor
    /// expression strings such as "10 mm" or "Width / 2".
    /// </summary>
    [IsVisibleInDynamoLibrary(true)]
    public static class InventorParameter
    {
        private static Inventor.ComponentDefinition ActiveDefinition =>
            PersistenceManager.IoC.GetInstance<IDocumentManager>().ActiveComponentDefinition;

        private static Inventor.Parameters ActiveParameters => ActiveDefinition.GetParameters();

        private static Document ActiveDocument => ActiveDefinition.GetDocument();

        private static Inventor.Parameter Find(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Parameter name is required.", nameof(name));
            try
            {
                return ActiveParameters[name];
            }
            catch (Exception ex)
            {
                throw new ArgumentException("No parameter named '" + name + "' in the active document.", nameof(name), ex);
            }
        }

        /// <summary>Names of every parameter (model, user and reference) in the active part or assembly.</summary>
        public static IList<string> Names()
        {
            return ActiveParameters.Cast<Inventor.Parameter>().Select(p => p.Name).ToList();
        }

        /// <summary>Names of the user parameters only.</summary>
        public static IList<string> UserParameterNames()
        {
            return ActiveParameters.UserParameters.Cast<Inventor.UserParameter>().Select(p => p.Name).ToList();
        }

        /// <summary>The parameter's current value, converted to its display units (e.g. mm, deg).</summary>
        public static double Value(string name)
        {
            var parameter = Find(name);
            return ConvertFromInternal(parameter, parameter.get_Units());
        }

        /// <summary>The parameter's expression string, e.g. "10 mm" or "Width / 2".</summary>
        public static string Expression(string name) => Find(name).Expression;

        /// <summary>The parameter's units, e.g. "mm", "deg", "ul".</summary>
        public static string Units(string name) => Find(name).get_Units();

        /// <summary>
        /// Sets a parameter's value in its own units and updates the document so dependent geometry
        /// rebuilds. Returns the value read back.
        /// </summary>
        public static double SetValue(string name, double value)
        {
            var parameter = Find(name);
            parameter.Expression = FormatExpression(value, parameter.get_Units());
            Update();
            return ConvertFromInternal(parameter, parameter.get_Units());
        }

        /// <summary>Sets a parameter's expression (e.g. "Width / 2 + 5 mm") and updates the document.</summary>
        public static string SetExpression(string name, string expression)
        {
            var parameter = Find(name);
            parameter.Expression = expression;
            Update();
            return parameter.Expression;
        }

        /// <summary>
        /// Adds a user parameter (or sets it when one of that name already exists) and updates the document.
        /// </summary>
        /// <param name="name">Parameter name, e.g. "Width".</param>
        /// <param name="value">Value in <paramref name="units"/>.</param>
        /// <param name="units">Inventor unit string, e.g. "mm", "in", "deg", or "ul" for unitless.</param>
        public static string AddUserParameter(string name, double value, string units = "mm")
        {
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Parameter name is required.", nameof(name));
            var parameters = ActiveParameters;
            var existing = parameters.UserParameters.Cast<Inventor.UserParameter>().FirstOrDefault(p => p.Name == name);
            if (existing != null)
            {
                existing.Expression = FormatExpression(value, units);
            }
            else
            {
                parameters.UserParameters.AddByExpression(name, FormatExpression(value, units), units);
            }
            Update();
            return name;
        }

        private static void Update()
        {
            ActiveDocument.Update();
        }

        private static string FormatExpression(double value, string units)
        {
            var number = value.ToString("R", System.Globalization.CultureInfo.InvariantCulture);
            return string.IsNullOrEmpty(units) || units == "ul" ? number : number + " " + units;
        }

        /// <summary>Parameter.Value is in Inventor's internal units (cm, radians); convert to the parameter's display units.</summary>
        private static double ConvertFromInternal(Inventor.Parameter parameter, string units)
        {
            var internalValue = Convert.ToDouble(parameter.Value, System.Globalization.CultureInfo.InvariantCulture);
            try
            {
                var uom = ActiveDocument.UnitsOfMeasure;
                return uom.ConvertUnits(internalValue, uom.GetDatabaseUnitsFromExpression(parameter.Expression, units), units);
            }
            catch
            {
                return internalValue;
            }
        }
    }
}
