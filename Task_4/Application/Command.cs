using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using System;
using Autodesk.Revit.UI.Selection;
using System.Collections.Generic;
using System.Linq;

namespace Task4
{
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        double studTickness = 0.15;
        double studSpacing = 2.0;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDocument = commandData.Application.ActiveUIDocument;
            Document document = uiDocument.Document;

            try
            {
                Reference wallRef = uiDocument.Selection.PickObject(ObjectType.Element, new SelectionFilter("Walls"), "Pick a Wall");
                Wall wall = document.GetElement(wallRef) as Wall;

                Solid wallSolid = GetWallSolid(wall);
                if (wallSolid == null)
                {
                    message = "Failed to get wall geometry.";
                    return Result.Failed;
                }

                // wall face
                Face wallFace = wallSolid.Faces.Cast<Face>().OrderByDescending(f => f.Area).FirstOrDefault();
                XYZ wallNormal = wallFace.ComputeNormal(new UV(0.5, 0.5));
                IList<CurveLoop> curveLoops = wallFace.GetEdgesAsCurveLoops();

                using (Transaction tr = new Transaction(document))
                {
                    tr.Start("Framing Wall");

                    CreateFraming(document, wall, wallSolid, wallFace, wallNormal, curveLoops);

                    tr.Commit();
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                TaskDialog.Show("Error", "An error occurred: " + ex.Message);
                return Result.Failed;
            }
        }

        #region Helper Methods

        private void CreateFraming(Document doc, Wall wall, Solid wallSolid, Face wallFace, XYZ wallNormal, IList<CurveLoop> curveLoops)
        {
            LocationCurve locCurve = wall.Location as LocationCurve;
            Curve wallCurve = locCurve.Curve;
            XYZ wallDir = (wallCurve.GetEndPoint(1) - wallCurve.GetEndPoint(0)).Normalize();
            double wallWidth = wall.Width;

            // vertical studs
            double spacing = UnitUtils.ConvertToInternalUnits(studSpacing, UnitTypeId.Feet);
            int pointCount = (int)(wallCurve.Length / spacing);

            for (int i = 1; i <= pointCount; i++)
            {
                double dist = i * spacing;
                if (Math.Abs(dist - wallCurve.Length) < 0.01) continue; // skip last point

                XYZ pointOnWall = wallCurve.Evaluate(dist / wallCurve.Length, true);
                CreateVerticalStudAtPoint(doc, pointOnWall, wallSolid, wallNormal, wallDir, wallWidth);
            }

            // bottom stud
            Transform moveTransform = Transform.CreateTranslation(wallNormal * wallWidth * 0.5);
            Curve movedCurve = wallCurve.CreateTransformed(moveTransform);

            Transform heightTransform = Transform.CreateTranslation(XYZ.BasisZ * studTickness);
            Curve offsetCurve = movedCurve.CreateTransformed(heightTransform);

            CreateModelCurve(doc, movedCurve, wallNormal, movedCurve.GetEndPoint(0));
            CreateModelCurve(doc, offsetCurve, wallNormal, offsetCurve.GetEndPoint(0));

            // uter studs
            CurveLoop outerLoop = curveLoops[0];
            foreach (Curve curve in outerLoop)
            {
                if (IsBottomEdge(curve)) continue; // skip bottom edge

                CreateStudPair(doc, curve, wallNormal);
            }

            // openings
            if (curveLoops.Count > 1)
            {
                IEnumerable<FamilyInstance> openings = GetWallOpenings(doc, wall);

                for (int i = 1; i < curveLoops.Count; ++i)
                {
                    CurveLoop openingLoop = curveLoops[i];
                    CreateOpeningFraming(doc, openingLoop, wallNormal,
                        openings.Any(o => o.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Doors));
                }
            }
        }

        private void CreateVerticalStudAtPoint(Document doc, XYZ pointOnWall, Solid wallSolid, XYZ wallNormal, XYZ wallDir, double wallWidth)
        {
            XYZ bottom = pointOnWall - XYZ.BasisZ * 100;
            XYZ top = pointOnWall + XYZ.BasisZ * 100;
            Line verticalLine = Line.CreateBound(bottom, top);

            // intersection with wall
            SolidCurveIntersection intersection = wallSolid.IntersectWithCurve(verticalLine, new SolidCurveIntersectionOptions());

            if (intersection.SegmentCount > 0)
            {
                for (int j = 0; j < intersection.SegmentCount; j++)
                {
                    Curve seg = intersection.GetCurveSegment(j);
                    if (seg == null) continue;

                    // trim and offset curve
                    Curve trimmedCurve = TrimCurve(seg, studTickness);
                    Transform moveTransform = Transform.CreateTranslation(wallNormal * wallWidth * 0.5);
                    Curve movedCurve = trimmedCurve.CreateTransformed(moveTransform);

                    // two sides of the stud
                    Transform t1 = Transform.CreateTranslation(wallDir * studTickness * 0.5);
                    Transform t2 = Transform.CreateTranslation(wallDir.Negate() * studTickness * 0.5);

                    Curve curve1 = movedCurve.CreateTransformed(t1);
                    Curve curve2 = movedCurve.CreateTransformed(t2);

                    CreateModelCurve(doc, curve1, wallNormal, curve1.GetEndPoint(0));
                    CreateModelCurve(doc, curve2, wallNormal, curve2.GetEndPoint(0));
                }
            }
        }

        private void CreateStudPair(Document doc, Curve curve, XYZ wallNormal)
        {
            XYZ p1 = curve.GetEndPoint(0);
            XYZ p2 = curve.GetEndPoint(1);

            XYZ lineDir = (p2 - p1).Normalize();
            XYZ offsetDir = wallNormal.CrossProduct(lineDir).Normalize();

            // main line
            CreateModelCurve(doc, Line.CreateBound(p1, p2), wallNormal, p1);

            // offset line
            XYZ p1Offset = p1 + offsetDir * studTickness;
            XYZ p2Offset = p2 + offsetDir * studTickness;
            CreateModelCurve(doc, Line.CreateBound(p1Offset, p2Offset), wallNormal, p1Offset);
        }

        private void CreateOpeningFraming(Document doc, CurveLoop openingLoop, XYZ wallNormal, bool isDoor)
        {
            List<Curve> curves = openingLoop.ToList();
            int skipIndex = -1;

            // doors, skip bottom edge
            if (isDoor)
            {
                double lowestZ = double.MaxValue;
                for (int j = 0; j < curves.Count; j++)
                {
                    Curve curve = curves[j];
                    double avgZ = (curve.GetEndPoint(0).Z + curve.GetEndPoint(1).Z) / 2.0;
                    if (avgZ < lowestZ)
                    {
                        lowestZ = avgZ;
                        skipIndex = j;
                    }
                }
            }

            for (int j = 0; j < curves.Count; j++)
            {
                // skip 
                if (isDoor && j == skipIndex) continue;

                // skip
                Curve curve = curves[j];
                if (IsBottomEdge(curve)) continue;

                CreateStudPair(doc, curve, wallNormal);
            }
        }

        private void CreateModelCurve(Document doc, Curve curve, XYZ normal, XYZ origin)
        {
            Plane plane = Plane.CreateByNormalAndOrigin(normal, origin);
            SketchPlane sketchPlane = SketchPlane.Create(doc, plane);
            doc.Create.NewModelCurve(curve, sketchPlane);
        }

        private bool IsBottomEdge(Curve curve)
        {
            XYZ p1 = curve.GetEndPoint(0);
            XYZ p2 = curve.GetEndPoint(1);

            if (Math.Abs(p1.Z - p2.Z) < 0.01)
            {
                double avgZ = (p1.Z + p2.Z) / 2.0;
                return avgZ < 1;
            }
            return false;
        }

        private List<FamilyInstance> GetWallOpenings(Document doc, Wall wall)
        {
            ElementClassFilter filter = new ElementClassFilter(typeof(FamilyInstance));
            FilteredElementCollector collector = new FilteredElementCollector(doc)
                .WherePasses(filter);

            BuiltInParameter hostParamId = BuiltInParameter.HOST_ID_PARAM;
            ElementId wallId = wall.Id;

            return collector
                .Cast<FamilyInstance>()
                .Where(instance =>
                {
                    Parameter hostParam = instance.get_Parameter(hostParamId);
                    return hostParam != null && hostParam.AsElementId().Equals(wallId);
                }).ToList();
        }

        private Solid GetWallSolid(Wall wall)
        {
            Options options = new Options();
            options.ComputeReferences = true;
            GeometryElement wallGeometry = wall.get_Geometry(options);

            return wallGeometry
                .Cast<GeometryObject>()
                .OfType<Solid>()
                .FirstOrDefault(solid => solid.Volume > 0);
        }

        private Curve TrimCurve(Curve curve, double trimLength)
        {
            if (curve.Length < 2 * trimLength)
                return curve;

            XYZ start = curve.GetEndPoint(0);
            XYZ end = curve.GetEndPoint(1);
            XYZ direction = (end - start).Normalize();

            XYZ newStart = start + direction * trimLength;
            XYZ newEnd = end - direction * trimLength;

            return Line.CreateBound(newStart, newEnd);
        }

        #endregion
    }
}