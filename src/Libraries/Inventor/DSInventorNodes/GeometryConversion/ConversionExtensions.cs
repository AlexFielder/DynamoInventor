using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.DesignScript.Runtime;
using Inventor;

using InventorServices;

namespace InventorLibrary.GeometryConversion
{
    [Browsable(false)]
    [IsVisibleInDynamoLibrary(false)]
    public static class ConversionExtensions
    {
        // Dynamo geometry is unitless; the Inventor API works in centimetres. Numbers pass straight
        // through, so a Dynamo unit is an Inventor centimetre.

        private static TransientGeometry TransientGeometry =>
            InventorServices.Persistence.PersistenceManager.InventorApplication.TransientGeometry;

        #region Proto -> Inventor types
        public static Inventor.Point ToPoint(this Autodesk.DesignScript.Geometry.Point xyz)
        {
            return TransientGeometry.CreatePoint(xyz.X, xyz.Y, xyz.Z);
        }

        public static Inventor.Vector ToVector(this Autodesk.DesignScript.Geometry.Vector v)
        {
            return TransientGeometry.CreateVector(v.X, v.Y, v.Z);
        }

        public static Inventor.UnitVector ToUnitVector(this Autodesk.DesignScript.Geometry.Vector v)
        {
            return TransientGeometry.CreateUnitVector(v.X, v.Y, v.Z);
        }
        #endregion

        #region Inventor -> Proto types
        public static Autodesk.DesignScript.Geometry.Point ToPoint(this Inventor.Point xyz)
        {
            return Autodesk.DesignScript.Geometry.Point.ByCoordinates(xyz.X, xyz.Y, xyz.Z);
        }

        public static Autodesk.DesignScript.Geometry.Vector ToVector(this Inventor.Vector v)
        {
            return Autodesk.DesignScript.Geometry.Vector.ByCoordinates(v.X, v.Y, v.Z);
        }

        public static Autodesk.DesignScript.Geometry.Vector ToVector(this Inventor.UnitVector v)
        {
            return Autodesk.DesignScript.Geometry.Vector.ByCoordinates(v.X, v.Y, v.Z);
        }
        #endregion
    }
}
