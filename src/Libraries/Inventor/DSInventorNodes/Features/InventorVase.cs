using System;
using Autodesk.DesignScript.Runtime;
using Inventor;
using InventorServices.Persistence;
using InventorServices.Utilities;
using DSCircle = Autodesk.DesignScript.Geometry.Circle;
using DSCurve = Autodesk.DesignScript.Geometry.Curve;

namespace InventorLibrary.Features
{
    // Note: the Inventor PIA declares PartDocument/_Document and SketchPoint/SketchEntity as unrelated
    // interfaces, so the explicit casts below are required; they are QueryInterface calls at runtime.
    /// <summary>
    /// Native Inventor output for the four-section Dynamo vase sample.
    /// Coordinates and radii are centimetres, matching DynamoInventor's geometry convention.
    /// </summary>
    [IsVisibleInDynamoLibrary(true)]
    public static class InventorVase
    {
        private const string OwnerSet = "DynamoInventorVase";
        private const string SchemaVersion = "1";
        private const double Tolerance = 1e-7;

        private sealed class Section
        {
            public double Radius;
            public double Z;
        }

        /// <summary>
        /// Creates or updates a native surface loft from four horizontal circles centred on the Z axis.
        /// The active document must be a part (.ipt). The same instance name updates the same vase in
        /// that part, including after saving/reopening it. A different name creates a separate vase.
        /// This sample deliberately makes an open surface, with no cap, wall thickness or solid body.
        /// </summary>
        /// <param name="crossSections">Exactly four circles, ordered from bottom to top, in centimetres.</param>
        /// <param name="instanceName">Stable key for this vase within the active part.</param>
        /// <returns>A status message naming the native loft feature and its document.</returns>
        public static string ByCircles(DSCurve[] crossSections, string instanceName = "DynamoVase")
        {
            Transaction transaction = null;
            var clock = System.Diagnostics.Stopwatch.StartNew();
            try
            {
                // Validate every input before changing any Inventor document.
                var sections = ReadSections(crossSections);
                DiagnosticLog.Write("InventorVase.ByCircles: sections read at " + clock.ElapsedMilliseconds + " ms");
                if (string.IsNullOrWhiteSpace(instanceName))
                    throw new ArgumentException("A non-empty vase instance name is required.", nameof(instanceName));

                var application = PersistenceManager.InventorApplication;
                var document = application.ActiveDocument as PartDocument;
                if (document == null)
                    throw new InvalidOperationException(
                        "Open a standard Inventor part (.ipt), make it active, then run the vase graph. " +
                        "This node does not create a part automatically or operate on an assembly/drawing.");

                var definition = document.ComponentDefinition;
                var feature = FindVase(definition, instanceName);
                var creating = feature == null;
                DiagnosticLog.Write("InventorVase.ByCircles: " + (creating ? "creating" : "updating") + " at " + clock.ElapsedMilliseconds + " ms");
                transaction = application.TransactionManager.StartTransaction((_Document)document,
                    (creating ? "Create " : "Update ") + instanceName);

                if (creating)
                {
                    var prefix = "DV_" + Guid.NewGuid().ToString("N");
                    feature = CreateVase(application, definition, sections, instanceName, prefix);
                }
                else
                {
                    var metadata = feature.AttributeSets[OwnerSet];
                    if (!string.Equals(Convert.ToString(metadata["Schema"].Value), SchemaVersion,
                        StringComparison.Ordinal))
                        throw new InvalidOperationException("This vase was created by a different node schema.");

                    var prefix = Convert.ToString(metadata["ParameterPrefix"].Value);
                    if (string.IsNullOrEmpty(prefix))
                        throw new InvalidOperationException("The vase's parameter metadata is missing.");

                    // Resolve all eight parameters before assigning any of them. No geometry is deleted
                    // or recreated: the existing planes, sketches and loft remain in the feature tree.
                    var parameters = new UserParameter[8];
                    for (var i = 0; i < 4; i++)
                    {
                        parameters[2 * i] = definition.Parameters.UserParameters[RadiusName(prefix, i)];
                        parameters[2 * i + 1] = definition.Parameters.UserParameters[HeightName(prefix, i)];
                    }
                    for (var i = 0; i < 4; i++)
                    {
                        // Inventor Parameter.Value uses database units: centimetres for lengths,
                        // regardless of whether the document displays mm, cm or inches.
                        parameters[2 * i].Value = sections[i].Radius;
                        parameters[2 * i + 1].Value = sections[i].Z;
                    }
                }

                DiagnosticLog.Write("InventorVase.ByCircles: geometry/parameters done at " + clock.ElapsedMilliseconds + " ms; updating document");
                if (!document.Update2(false))
                    throw new InvalidOperationException(
                        "Inventor could not regenerate the vase. The transaction will be rolled back. " +
                        "Check the section sizes and any existing model errors.");

                // Read the status before committing: any failure above/below still rolls back as a unit.
                var result = (creating ? "Created " : "Updated ") + feature.Name + " in " +
                    document.DisplayName + " (4-section surface loft; coordinates in cm).";
                transaction.End();
                transaction = null;
                DiagnosticLog.Write("InventorVase.ByCircles: " + result + " (" + clock.ElapsedMilliseconds + " ms)");
                return result;
            }
            catch (Exception ex)
            {
                if (transaction != null)
                {
                    try { transaction.Abort(); }
                    catch (Exception rollbackError)
                    {
                        DiagnosticLog.Write("InventorVase.ByCircles rollback", rollbackError);
                    }
                }
                DiagnosticLog.Write("InventorVase.ByCircles", ex);
                throw;
            }
        }

        private static Section[] ReadSections(DSCurve[] curves)
        {
            if (curves == null || curves.Length != 4)
                throw new ArgumentException("The vase needs exactly four circular sections.", nameof(curves));

            var result = new Section[4];
            for (var i = 0; i < curves.Length; i++)
            {
                var circle = curves[i] as DSCircle;
                if (circle == null)
                    throw new ArgumentException("Section " + (i + 1) + " must be a Dynamo Circle.");

                // Do not dispose the input circles: they are also used by Dynamo's Surface.ByLoft.
                using (var center = circle.CenterPoint)
                using (var normal = circle.Normal)
                {
                    var radius = circle.Radius;
                    if (!Finite(radius) || radius <= Tolerance || !Finite(center.X) ||
                        !Finite(center.Y) || !Finite(center.Z))
                        throw new ArgumentException("Circle coordinates must be finite and radii must be positive.");
                    if (Math.Abs(center.X) > Tolerance || Math.Abs(center.Y) > Tolerance)
                        throw new ArgumentException("This vase sample requires all circles to be centred on the Z axis.");

                    var length = Math.Sqrt(normal.X * normal.X + normal.Y * normal.Y + normal.Z * normal.Z);
                    if (!Finite(length) || length <= Tolerance ||
                        Math.Abs(normal.X / length) > Tolerance || Math.Abs(normal.Y / length) > Tolerance)
                        throw new ArgumentException("All vase circles must be parallel to the XY plane.");
                    if (i > 0 && center.Z - result[i - 1].Z <= Tolerance)
                        throw new ArgumentException("Order the circles from bottom to top with distinct, increasing heights.");

                    result[i] = new Section { Radius = radius, Z = center.Z };
                }
            }
            return result;
        }

        private static LoftFeature FindVase(PartComponentDefinition definition, string instanceName)
        {
            LoftFeature result = null;
            foreach (LoftFeature candidate in definition.Features.LoftFeatures)
            {
                if (!candidate.AttributeSets.get_NameIsUsed(OwnerSet)) continue;
                var metadata = candidate.AttributeSets[OwnerSet];
                if (!string.Equals(Convert.ToString(metadata["Instance"].Value), instanceName,
                    StringComparison.Ordinal)) continue;

                if (result != null)
                    throw new InvalidOperationException("More than one loft is tagged as '" + instanceName +
                        "'. Resolve the duplicate model features before running this graph.");
                result = candidate;
            }
            return result;
        }

        private static LoftFeature CreateVase(Inventor.Application application,
            PartComponentDefinition definition, Section[] sections, string instanceName, string prefix)
        {
            var profiles = application.TransientObjects.CreateObjectCollection();
            var geometry = application.TransientGeometry;
            var sketches = new PlanarSketch[4];
            var planes = new WorkPlane[4];
            var xyPlane = definition.WorkPlanes[3];

            for (var i = 0; i < 4; i++)
            {
                var radiusName = RadiusName(prefix, i);
                var heightName = HeightName(prefix, i);
                definition.Parameters.UserParameters.AddByValue(radiusName, sections[i].Radius,
                    UnitsTypeEnum.kCentimeterLengthUnits);
                definition.Parameters.UserParameters.AddByValue(heightName, sections[i].Z,
                    UnitsTypeEnum.kCentimeterLengthUnits);

                // An offset plane is created even at z=0, so every section remains independently editable.
                var plane = definition.WorkPlanes.AddByPlaneAndOffset(xyPlane, heightName, false);
                plane.Name = instanceName + "_Plane" + (i + 1);
                planes[i] = plane;

                var sketch = definition.Sketches.Add(plane);
                sketch.Name = instanceName + "_Section" + (i + 1);
                sketches[i] = sketch;

                var center = sketch.SketchPoints.Add(geometry.CreatePoint2d(0, 0), false);
                sketch.GeometricConstraints.AddGround((SketchEntity)center);
                var circle = sketch.SketchCircles.AddByCenterRadius(center, sections[i].Radius);
                var radius = sketch.DimensionConstraints.AddRadius((SketchEntity)circle,
                    geometry.CreatePoint2d(sections[i].Radius * 1.3, sections[i].Radius * 0.3), false);
                radius.Parameter.Expression = radiusName;
                profiles.Add(sketch.Profiles.AddForSurface(circle));
            }

            var loftDefinition = definition.Features.LoftFeatures.CreateLoftDefinition(profiles,
                PartFeatureOperationEnum.kSurfaceOperation);
            loftDefinition.Closed = false;
            loftDefinition.MergeTangentFaces = true;
            var feature = definition.Features.LoftFeatures.Add(loftDefinition);
            feature.Name = instanceName;

            var metadata = feature.AttributeSets.Add(OwnerSet);
            metadata.Add("Schema", ValueTypeEnum.kStringType, SchemaVersion);
            metadata.Add("Instance", ValueTypeEnum.kStringType, instanceName);
            metadata.Add("ParameterPrefix", ValueTypeEnum.kStringType, prefix);

            for (var i = 0; i < 4; i++)
            {
                sketches[i].Visible = false;
                planes[i].Visible = false;
            }
            return feature;
        }

        private static string RadiusName(string prefix, int index) => prefix + "_Radius" + (index + 1);
        private static string HeightName(string prefix, int index) => prefix + "_Z" + (index + 1);
        private static bool Finite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
    }
}
