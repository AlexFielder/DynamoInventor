using System;
using Autodesk.DesignScript.Geometry;
using Autodesk.DesignScript.Runtime;
using InventorLibrary.GeometryConversion;
using InventorServices.Persistence;
using InventorServices.Utilities;

namespace InventorLibrary.WorkGeometry
{
    /// <summary>
    /// A fixed Inventor work plane created and re-bound by Dynamo (trace-tracked across re-runs), in
    /// the active part or assembly. Fixed planes are the one definition type both document types share.
    /// </summary>
    [IsVisibleInDynamoLibrary(true)]
    public class InventorWorkPlane
    {
        private readonly IObjectBinder binder;

        internal Inventor.WorkPlane InternalWorkPlane { get; private set; }

        [IsVisibleInDynamoLibrary(false)]
        private InventorWorkPlane(Point origin, Vector xAxis, Vector yAxis, IObjectBinder binder)
        {
            this.binder = binder;
            var documents = binder.DocumentManager;
            binder.ContextManager.BindingContextManager = documents.ActiveReferenceKeyManager;

            Inventor.WorkPlane existing;
            if (binder.GetObjectFromTrace<Inventor.WorkPlane>(out existing))
            {
                existing.SetFixed(origin.ToPoint(), xAxis.ToUnitVector(), yAxis.ToUnitVector());
                InternalWorkPlane = existing;
            }
            else
            {
                InternalWorkPlane = documents.ActiveComponentDefinition.GetWorkPlanes().AddFixed(
                    origin.ToPoint(), xAxis.ToUnitVector(), yAxis.ToUnitVector(), false);
                binder.SetObjectForTrace<InventorWorkPlane>(InternalWorkPlane);
            }
        }

        /// <summary>Creates (or on re-run, repositions) a fixed work plane from an origin and two in-plane directions.</summary>
        /// <param name="origin">Plane origin, in Inventor's internal centimetres.</param>
        /// <param name="xAxis">Direction of the plane's X axis.</param>
        /// <param name="yAxis">Direction of the plane's Y axis; must not be parallel to xAxis.</param>
        public static InventorWorkPlane ByOriginXAxisYAxis(Point origin, Vector xAxis, Vector yAxis)
        {
            if (origin == null) throw new ArgumentNullException(nameof(origin));
            if (xAxis == null) throw new ArgumentNullException(nameof(xAxis));
            if (yAxis == null) throw new ArgumentNullException(nameof(yAxis));
            try
            {
                var binder = PersistenceManager.IoC.GetInstance<IObjectBinder>();
                return new InventorWorkPlane(origin, xAxis, yAxis, binder);
            }
            catch (Exception ex)
            {
                InventorServices.Utilities.DiagnosticLog.Write("InventorWorkPlane.ByOriginXAxisYAxis", ex);
                throw;
            }
        }

        /// <summary>Creates (or on re-run, repositions) a fixed work plane coincident with a Dynamo plane.</summary>
        public static InventorWorkPlane ByPlane(Plane plane)
        {
            if (plane == null) throw new ArgumentNullException(nameof(plane));
            return ByOriginXAxisYAxis(plane.Origin, plane.XAxis, plane.YAxis);
        }

        /// <summary>The work plane's name in the Inventor browser, e.g. "Work Plane1".</summary>
        public string Name => InternalWorkPlane.Name;

        /// <summary>The work plane as a Dynamo plane (origin and normal read back from Inventor).</summary>
        public Plane Plane
        {
            get
            {
                var p = InternalWorkPlane.Plane;
                return Plane.ByOriginNormal(p.RootPoint.ToPoint(), p.Normal.ToVector());
            }
        }

        public override string ToString() => "InventorWorkPlane(" + Name + ")";
    }
}
