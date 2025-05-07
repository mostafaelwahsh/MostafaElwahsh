using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Task2.Application;
using Autodesk.Revit.UI;

namespace Task2.Utilities
{
    public static class RoomHelper
    {

        public static Room GetRoomsNextToSelectedWall(Wall selectedWall, out IList<Autodesk.Revit.DB.BoundarySegment> boundarySegment)
        {
            boundarySegment = null;
            Room roomsFound = null;

            var bathroomRooms = new FilteredElementCollector(Command.CommandDoc)
                          .OfCategory(BuiltInCategory.OST_Rooms)
                          .WhereElementIsNotElementType()
                          .Where(room => room.Name.Contains("Bathroom"))
                          .Cast<Room>()
                          .ToList();

            if (!bathroomRooms.Any())
            {
                TaskDialog.Show("Alert", "No Bathrooms Found!");
                return null;
            }

            SpatialElementBoundaryOptions options = new SpatialElementBoundaryOptions
            {
                SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Finish
            };

            foreach (Room room in bathroomRooms)
            {
                foreach (IList<Autodesk.Revit.DB.BoundarySegment> boundSegList in room.GetBoundarySegments(options))
                {
                    foreach (Autodesk.Revit.DB.BoundarySegment boundSeg in boundSegList)
                    {
                        if (IsWallSegmentOfSelectedWall(boundSeg, selectedWall))
                        {
                            roomsFound = room;
                            boundarySegment = boundSegList;
                        }
                    }
                }
            }
            return roomsFound;
        }
        private static bool IsWallSegmentOfSelectedWall(BoundarySegment boundSeg, Wall selectedWall)
        {
            Element e = Command.CommandDoc.GetElement(boundSeg.ElementId);
            Wall wall = e as Wall;
            return wall != null && wall.Id == selectedWall.Id;
        }
        public static List<XYZ> GetDoorsLocationInRoom(Room room, IList<BoundarySegment> boundarySegmentList)
        {
            if (boundarySegmentList == null)
            {
                return null;
            }

            List<XYZ> doorsLocationPoint = new List<XYZ>();

            foreach (BoundarySegment boundSeg in boundarySegmentList)
            {
                if (boundSeg == null)
                {
                    continue;
                }

                var wallInRoom = Command.CommandDoc.GetElement(boundSeg.ElementId) as Wall;

                if (wallInRoom == null || !(wallInRoom is HostObject))
                {
                    continue;
                }

                var wallHostObj = wallInRoom as HostObject;
                var hostedElementsOnWall = wallHostObj.FindInserts(true, true, true, true);

                if (hostedElementsOnWall != null && hostedElementsOnWall.Any())
                {
                    var famInstanceCollector = new FilteredElementCollector(Command.CommandDoc, hostedElementsOnWall)
                        .OfCategory(BuiltInCategory.OST_Doors)
                        .WhereElementIsNotElementType()
                        .Cast<FamilyInstance>()
                        .Where(A => A != null && (A.ToRoom?.Name == room.Name || A.FromRoom?.Name == room.Name))
                        .ToList();

                    doorsLocationPoint.AddRange(famInstanceCollector.Select(famInstance => (famInstance.Location as LocationPoint)?.Point)
                        .Where(locationPoint => locationPoint != null));
                }
            }

            if (!doorsLocationPoint.Any())
            {
                TaskDialog.Show("Error", $"No doors in this room");
                return null;
            }

            return doorsLocationPoint;
        }
        public static List<Curve> GetDoorsCurvesInRoom(Room room, IList<BoundarySegment> boundarySegmentList)
        {
            if (boundarySegmentList == null)
            {
                return null;
            }

            List<Curve> doorsCurves = new List<Curve>();

            foreach (BoundarySegment boundSeg in boundarySegmentList)
            {
                if (boundSeg == null)
                {
                    continue;
                }

                var wallInRoom = Command.CommandDoc.GetElement(boundSeg.ElementId) as Wall;

                if (wallInRoom == null || !(wallInRoom is HostObject))
                {
                    continue;
                }

                var wallHostObj = wallInRoom as HostObject;
                var hostedElementsOnWall = wallHostObj.FindInserts(true, true, true, true);

                if (hostedElementsOnWall != null && hostedElementsOnWall.Any())
                {
                    var famInstanceCollector = new FilteredElementCollector(Command.CommandDoc, hostedElementsOnWall)
                        .OfCategory(BuiltInCategory.OST_Doors)
                        .WhereElementIsNotElementType()
                        .Cast<FamilyInstance>()
                        .Where(A => A != null && (A.ToRoom?.Name == room.Name || A.FromRoom?.Name == room.Name))
                        .ToList();

                    foreach (var famInstance in famInstanceCollector)
                    {
                        var geometryElement = famInstance.get_Geometry(new Options());
                        foreach (var geometryInstance in geometryElement)
                        {
                            var instanceGeometry = geometryInstance as GeometryInstance;
                            if (instanceGeometry != null)
                            {
                                foreach (var geometryObject in instanceGeometry.GetInstanceGeometry())
                                {
                                    var curve = geometryObject as Curve;
                                    if (curve != null)
                                    {
                                        doorsCurves.Add(curve);
                                    }
                                }
                            }
                        }
                    }
                }
            }

            if (!doorsCurves.Any())
            {
                TaskDialog.Show("Error", $"No doors in this room");
                return null;
            }

            return doorsCurves;
        }

    }
}
