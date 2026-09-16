using System;
using Autodesk.DesignScript.Runtime;
using InventorLibrary.GeometryConversion;
using InventorServices.Persistence;
using InventorServices.Utilities;
using Point = Autodesk.DesignScript.Geometry.Point;

namespace InventorLibrary.WorkGeometry
{
    /// <summary>
    /// A fixed Inventor work point created and re-bound by Dynamo (trace-tracked across re-runs), in
    /// the active part or assembly. Assemblies only allow fixed work points; parts allow every
    /// definition type, but a fixed point is what a Dynamo-driven coordinate maps to in both.
    /// Named to avoid clashing with the generated API wrapper InventorLibrary.API.InvWorkPoint and
    /// with Inventor's own WorkPoint interop type, both of which DesignScript can see.
    /// </summary>
    [IsVisibleInDynamoLibrary(true)]
    public class InventorWorkPoint
    {
        private readonly IObjectBinder binder;

        internal Inventor.WorkPoint InternalWorkPoint { get; private set; }

        [IsVisibleInDynamoLibrary(false)]
        private InventorWorkPoint(Point point, IObjectBinder binder)
        {
            this.binder = binder;
            var documents = binder.DocumentManager;

            // Bind reference keys against whichever document is active (part or assembly).
            binder.ContextManager.BindingContextManager = documents.ActiveReferenceKeyManager;

            Inventor.WorkPoint existing;
            if (binder.GetObjectFromTrace<Inventor.WorkPoint>(out existing))
            {
                // Same node, re-run: move the point we made last time instead of making another.
                existing.SetFixed(point.ToPoint());
                InternalWorkPoint = existing;
            }
            else
            {
                InternalWorkPoint = documents.ActiveComponentDefinition.GetWorkPoints().AddFixed(point.ToPoint(), false);
                binder.SetObjectForTrace<InventorWorkPoint>(InternalWorkPoint);
            }
        }

        /// <summary>Creates (or on re-run, moves) a fixed work point at the given point in the active part or assembly.</summary>
        /// <param name="point">Location, in Inventor's internal centimetres.</param>
        public static InventorWorkPoint ByPoint(Point point)
        {
            if (point == null) throw new ArgumentNullException(nameof(point));
            try
            {
                var binder = PersistenceManager.IoC.GetInstance<IObjectBinder>();
                return new InventorWorkPoint(point, binder);
            }
            catch (Exception ex)
            {
                InventorServices.Utilities.DiagnosticLog.Write("InventorWorkPoint.ByPoint", ex);
                throw;
            }
        }

        /// <summary>The work point's name in the Inventor browser, e.g. "Work Point1".</summary>
        public string Name => InternalWorkPoint.Name;

        /// <summary>The work point's current location as a Dynamo point.</summary>
        public Point Point => InternalWorkPoint.Point.ToPoint();

        public override string ToString() => "InventorWorkPoint(" + Name + ")";
    }
}
