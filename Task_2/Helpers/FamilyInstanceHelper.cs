using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using Task_2.Helpers;

namespace Task_2.Helpers
{
    public class FamilyInstanceHelper
    {
        public static void MoveFamilyToCorrectPlace(FamilyInstance familyInstanceCreated, Curve wallCurve, IList<BoundarySegment> boundarySegment, Document doc)
        {
            var listOfWallCurves = boundarySegment.Select(boundSeg =>
            {
                Element e = doc.GetElement(boundSeg.ElementId);
                Wall wall = e as Wall;
                return (wall.Location as LocationCurve).Curve;
            }).ToList();

            var famInstanceCurves = FamilyInstanceBoundaryLines(5, 5, (familyInstanceCreated.Location as LocationPoint).Point);


            if (GeometryHelper.AreLinesIntersecting(listOfWallCurves, famInstanceCurves))
            {
                // Determine the direction to move based on the closest point
                XYZ closestPointOnWallCurve = GeometryHelper.GetClosestPointOnCurve(wallCurve, (familyInstanceCreated.Location as LocationPoint).Point);

                // Calculate the translation vector
                XYZ translationVector = GeometryHelper.GetTranslationVector((familyInstanceCreated.Location as LocationPoint).Point, closestPointOnWallCurve);

                // Move the family instance
                ElementTransformUtils.MoveElement(doc, familyInstanceCreated.Id, translationVector);

                // Check for intersection after the move
                famInstanceCurves = FamilyInstanceBoundaryLines(5, 5, (familyInstanceCreated.Location as LocationPoint).Point);
                while (GeometryHelper.AreLinesIntersecting(listOfWallCurves, famInstanceCurves))
                {
                    // Move further if still intersecting
                    translationVector = GeometryHelper.GetTranslationVector((familyInstanceCreated.Location as LocationPoint).Point, closestPointOnWallCurve);
                    ElementTransformUtils.MoveElement(doc, familyInstanceCreated.Id, translationVector);

                    // Check for intersection again
                    famInstanceCurves = FamilyInstanceBoundaryLines(5, 5, (familyInstanceCreated.Location as LocationPoint).Point);
                }
            }
        }
        public static List<Curve> FamilyInstanceBoundaryLines(double familyInsWidth, double familyInsHeight, XYZ familyInsLocPoint, bool IsVerticalWall = false)
        {
            if (IsVerticalWall)
            {
                var height = familyInsHeight;
                var width = familyInsWidth;
                familyInsHeight = width;
                familyInsWidth = height;
            }
            List<Curve> lines = new List<Curve>();
            var instanceLine1 = Line.CreateBound(new XYZ(familyInsLocPoint.X - (familyInsWidth / 2), familyInsLocPoint.Y, 0), new XYZ(familyInsLocPoint.X + (familyInsWidth / 2), familyInsLocPoint.Y, 0)) as Curve;
            var instanceLine2 = Line.CreateBound(new XYZ(familyInsLocPoint.X - (familyInsWidth / 2), familyInsLocPoint.Y + (familyInsHeight / 2), 0), new XYZ(familyInsLocPoint.X + (familyInsWidth / 2), familyInsLocPoint.Y + (familyInsHeight / 2), 0)) as Curve;
            var instanceLine3 = Line.CreateBound(new XYZ(familyInsLocPoint.X - (familyInsWidth / 2), familyInsLocPoint.Y - (familyInsHeight / 2), 0), new XYZ(familyInsLocPoint.X + (familyInsWidth / 2), familyInsLocPoint.Y - (familyInsHeight / 2), 0)) as Curve;

            lines.Add(instanceLine1);
            lines.Add(instanceLine2);
            lines.Add(instanceLine3);

            return lines;
        }
        public static void VerticalHandOrientationCriterea(FamilyInstance familyCreated, Curve curveWall, XYZ familyLocationPoint, XYZ roomMidPoint, Document doc)
        {
            XYZ directionToFamily = familyLocationPoint - roomMidPoint;

            if (Math.Abs(curveWall.GetEndPoint(0).X - curveWall.GetEndPoint(1).X) < 0.001)
            {
                // Check if the wall is left or right of the room midpoint
                if (curveWall.GetEndPoint(0).X < roomMidPoint.X)
                {
                    // Wall is left of room midpoint
                    if (directionToFamily.Y < 0)
                    {
                        // Below room centroid
                        if (!familyCreated.HandOrientation.IsAlmostEqualTo(new XYZ(0, 1, 0)))
                        {
                            familyCreated.flipHand();
                        }

                        var translation = new XYZ(0, 1.5, 0);
                        ElementTransformUtils.MoveElement(doc, familyCreated.Id, translation);
                    }
                    else
                    {
                        // Above room centroid
                        if (!familyCreated.HandOrientation.IsAlmostEqualTo(new XYZ(0, -1, 0)))
                        {
                            familyCreated.flipHand();
                        }
                        var translation = new XYZ(0, -1.5, 0);
                        ElementTransformUtils.MoveElement(doc, familyCreated.Id, translation);
                    }
                }
                else
                {
                    // Wall is right of room midpoint
                    if (directionToFamily.Y < 0)
                    {
                        // Below room centroid
                        if (!familyCreated.HandOrientation.IsAlmostEqualTo(new XYZ(0, 1, 0)))
                        {
                            familyCreated.flipHand();
                        }
                        var translation = new XYZ(0, 1.5, 0);
                        ElementTransformUtils.MoveElement(doc, familyCreated.Id, translation);
                    }
                    else
                    {
                        // Above room centroid
                        if (!familyCreated.HandOrientation.IsAlmostEqualTo(new XYZ(0, -1, 0)))
                        {
                            familyCreated.flipHand();
                        }
                        var translation = new XYZ(0, -1.5, 0);
                        ElementTransformUtils.MoveElement(doc, familyCreated.Id, translation);
                    }
                }
            }

        }
        public static void HorizontalHandOrientationCriteria(FamilyInstance familyCreated, Curve curveWall, XYZ familyLocationPoint, XYZ roomMidPoint, Document doc)
        {
            XYZ directionToFamily = familyLocationPoint - roomMidPoint;

            if (Math.Abs(curveWall.GetEndPoint(0).Y - curveWall.GetEndPoint(1).Y) < 0.001)
            {
                // Check if the wall is left or right of the room midpoint
                if (familyLocationPoint.X < roomMidPoint.X)
                {
                    // Wall is left of room midpoint
                    if (directionToFamily.Y < 0)
                    {
                        // Below room centroid
                        if (!familyCreated.HandOrientation.IsAlmostEqualTo(new XYZ(1, 0, 0)))
                        {
                            familyCreated.flipHand();
                        }
                        var translation = new XYZ(1.5, 0, 0);
                        ElementTransformUtils.MoveElement(doc, familyCreated.Id, translation);
                    }
                    else
                    {
                        // Above room centroid
                        if (!familyCreated.HandOrientation.IsAlmostEqualTo(new XYZ(1, 0, 0)))
                        {
                            familyCreated.flipHand();
                        }
                        var translation = new XYZ(1.5, 0, 0);
                        ElementTransformUtils.MoveElement(doc, familyCreated.Id, translation);
                    }
                }
                else
                {
                    // Wall is right of room midpoint
                    if (directionToFamily.Y < 0)
                    {
                        // Below room centroid
                        if (!familyCreated.HandOrientation.IsAlmostEqualTo(new XYZ(-1, 0, 0)))
                        {
                            familyCreated.flipHand();
                        }
                        var translation = new XYZ(-1.5, 0, 0);
                        ElementTransformUtils.MoveElement(doc, familyCreated.Id, translation);
                    }
                    else
                    {
                        // Above room centroid
                        if (!familyCreated.HandOrientation.IsAlmostEqualTo(new XYZ(-1, 0, 0)))
                        {
                            familyCreated.flipHand();
                        }
                        var translation = new XYZ(-1.5, 0, 0);
                        ElementTransformUtils.MoveElement(doc, familyCreated.Id, translation);
                    }
                }
            }
        }
        public static void FamilyOrientationHandler(Document doc, FamilyInstance familyCreated, XYZ familyOrientation, Curve selectedWallCurveInRoom, XYZ familyLocationPoint, XYZ roomMidPoint)
        {
            if (!familyCreated.FacingOrientation.IsAlmostEqualTo(familyOrientation))
            {
                familyCreated.flipFacing();
            }

            if (GeometryHelper.IsVerticalCurve(selectedWallCurveInRoom))
            {
                VerticalHandOrientationCriterea(familyCreated, selectedWallCurveInRoom, familyLocationPoint, roomMidPoint, doc);
            }
            else
            {
                HorizontalHandOrientationCriteria(familyCreated, selectedWallCurveInRoom, familyLocationPoint, roomMidPoint, doc);
            }
        }
    }
}
