using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Task2.Application;

namespace Task_2.Helpers
{
    public class GeometryHelper
    {
        public static XYZ FarthestPointToCurve(List<XYZ> locPointsList, Curve wallCurve)
        {

            XYZ doorEndPoint = locPointsList.First();
            IList<XYZ> wallEndPoints = wallCurve.Tessellate();

            double maxDistance = double.MinValue;
            XYZ farthestPoint = null;

            foreach (XYZ wallPoint in wallEndPoints)
            {
                double distance = wallPoint.DistanceTo(doorEndPoint);

                if (distance > maxDistance)
                {
                    maxDistance = distance;
                    farthestPoint = wallPoint;
                }
            }

            return farthestPoint;
        }

        public static bool IsVerticalCurve(Curve curve)
        {
            //For simplicity, here we assume it's vertical if the difference in X coordinates is negligible
            return Math.Abs(curve.GetEndPoint(0).X - curve.GetEndPoint(1).X) < 0.001;
        }
        public static Curve GetWallCurveInsideRoom(IList<BoundarySegment> boundarySegmentList, Wall wallElement)
        {
            var matchingSegment = boundarySegmentList.FirstOrDefault(boundSeg => boundSeg.ElementId == wallElement.Id);

            if (matchingSegment != null)
            {
                return matchingSegment.GetCurve();
            }
            return null;
        }
        public static XYZ GetCorrectFamilyOrientation(Room room, XYZ point, Curve wallCurve, out XYZ roomCentroid)
        {
            roomCentroid = (room.Location as LocationPoint).Point;
            XYZ directionVector = roomCentroid - point;

            if (IsVerticalCurve(wallCurve))
            {
                directionVector = new XYZ(Math.Sign(directionVector.X), 0, 0);
            }
            else
            {
                directionVector = new XYZ(0, Math.Sign(directionVector.Y), 0);
            }

            return directionVector;
        }
        public static bool AreLinesIntersecting(List<Curve> instanceLines1, List<Curve> instanceLines2)
        {
            foreach (Curve line1 in instanceLines1)
            {
                foreach (Curve line2 in instanceLines2)
                {
                    if (line1.Intersect(line2) == SetComparisonResult.Overlap)
                    {
                        return true;
                    }
                }
            }
            return false;
        }
        public static XYZ GetClosestPointOnCurve(Curve curve, XYZ point)
        {
            double param = curve.Project(point).Parameter;
            return curve.Evaluate(param, false);
        }
        public static XYZ GetTranslationVector(XYZ fromPoint, XYZ toPoint, double distance = 0.1)
        {
            // Calculate the translation vector from 'fromPoint' to 'toPoint' with a specified distance
            XYZ direction = toPoint - fromPoint;
            direction = direction.Normalize();
            return direction * distance;
        }
    }
}
