using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB.Architecture;
using System.Collections.Generic;
using System.Linq;
using Task_3.Helpers;

namespace Task_3.Application
{
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uIDocument = commandData.Application.ActiveUIDocument;
            Document document = uIDocument.Document;
            string floorTypeName = "Generic 150mm";

            try
            {
                if (!(document.ActiveView is ViewPlan viewPlan) || viewPlan.ViewType != ViewType.FloorPlan)
                {
                    TaskDialog.Show("Error", "run this command in a floor plan view.");
                    return Result.Failed;
                }

                // all rooms
                FilteredElementCollector allrooms = new FilteredElementCollector(document)
                    .OfCategory(BuiltInCategory.OST_Rooms)
                    .WhereElementIsNotElementType();

                List<Room> rooms = allrooms.Cast<Room>()
                    .Where(r => r.Location != null && r.Area > 0)
                    .ToList();

                if (rooms.Count == 0)
                {
                    TaskDialog.Show("Error", "no rooms found in the document.");
                    return Result.Failed;
                }

                using (Transaction tr = new Transaction(document))
                {
                    tr.Start("Rooms Thresholds");

                    foreach (Room room in rooms)
                    {
                        try
                        {
                            Level level = room.Level;

                            // assume floor thickness if no existing floor
                            double offset = 0;
                            double floorThickness = 0.15;

                            Floor outfloor;
                            Solid roomFloorSolid = GetOrCreateRoomFloor(document, room, level, out outfloor, out offset);

                            if (roomFloorSolid == null)
                                continue;


                            if (outfloor != null)
                            {
                                floorThickness = outfloor.get_Parameter(BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM).AsDouble();
                                floorTypeName = outfloor.FloorType.Name;
                            }

                            List<Doorthreshold> roomThresholds = GetRoomAllThresholds(document, room);

                            Solid lastSolid = roomFloorSolid;
                            foreach (Doorthreshold doorthreshold in roomThresholds)
                            {
                                Solid ThresholdSolid = GetThresholdSolid(doorthreshold, floorThickness, offset);

                                Solid solid = BooleanOperationsUtils.ExecuteBooleanOperation(lastSolid, ThresholdSolid,
                                    BooleanOperationsType.Union);

                                lastSolid = solid;
                            }

                            // new floor
                            Floor newFloor = CreateFloorFromSolid(document, lastSolid, floorTypeName, level);

                            if (newFloor == null)
                                continue;

                            // offset parameter
                            Parameter heightParameter = newFloor.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM);
                            if (heightParameter != null && !heightParameter.IsReadOnly)
                            {
                                heightParameter.Set(offset);
                            }

                            // parameters
                            if (outfloor != null)
                            {
                                CopyFloorParameters(outfloor, newFloor);
                                document.Delete(outfloor.Id);
                            }
                        }
                        catch (Exception ex)
                        { }
                    }

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

        private List<Doorthreshold> GetRoomAllThresholds(Document doc, Room room)
        {
            List<Doorthreshold> thresholds = new List<Doorthreshold>();

            BoundingBoxXYZ roombb = room.get_BoundingBox(null);
            Outline outline = new Outline(roombb.Min, roombb.Max);
            BoundingBoxIntersectsFilter boundingBoxIntersectsFilter = new BoundingBoxIntersectsFilter(outline);

            // all doors
            var doors = new FilteredElementCollector(doc)
                                          .OfCategory(BuiltInCategory.OST_Doors)
                                          .WherePasses(boundingBoxIntersectsFilter)
                                          .Cast<FamilyInstance>()
                                          .ToList();

            foreach (FamilyInstance d in doors)
            {
                if (d != null && d.Host != null)
                {
                    ElementId typeId = d.GetTypeId();
                    Element type = doc.GetElement(typeId);

                    Wall hostWall = doc.GetElement(d.Host.Id) as Wall;
                    if (hostWall == null) continue;

                    Doorthreshold threshold = new Doorthreshold();

                    threshold.Room = room;
                    threshold.Door = d;
                    threshold.HostWall = hostWall;
                    threshold.Locatin = (d.Location as LocationPoint).Point;
                    threshold.Width = type.get_Parameter(BuiltInParameter.DOOR_WIDTH).AsDouble();
                    threshold.Depth = hostWall.Width / 2.0;

                    thresholds.Add(threshold);
                }
            }

            return thresholds;
        }

        private Solid GetThresholdSolid(Doorthreshold doorthreshold, double floorThickness, double zOffset)
        {
            XYZ location = doorthreshold.Locatin;
            Wall wall = doorthreshold.HostWall;
            Curve curve = (wall.Location as LocationCurve).Curve;

            XYZ WallDirection = (curve.GetEndPoint(1) - curve.GetEndPoint(0)).Normalize();

            XYZ p1 = location + WallDirection * doorthreshold.Width / 2.0;
            XYZ p2 = location + WallDirection.Negate() * doorthreshold.Width / 2.0;

            XYZ FaceDirection = doorthreshold.Door.FacingOrientation;
            XYZ roomLocation = (doorthreshold.Room.Location as LocationPoint).Point;
            XYZ fromDoorToRoomCenter = (roomLocation - doorthreshold.Locatin).Normalize();

            XYZ p3;
            XYZ p4;
            if (FaceDirection.DotProduct(fromDoorToRoomCenter) > 0)
            {
                p3 = p2 + FaceDirection * doorthreshold.Depth;
                p4 = p1 + FaceDirection * doorthreshold.Depth;
            }
            else
            {
                p3 = p2 - FaceDirection * doorthreshold.Depth;
                p4 = p1 - FaceDirection * doorthreshold.Depth;
            }

            Line line1 = Line.CreateBound(p1, p2);
            Line line2 = Line.CreateBound(p2, p3);
            Line line3 = Line.CreateBound(p3, p4);
            Line line4 = Line.CreateBound(p4, p1);

            CurveLoop curves = new CurveLoop();
            curves.Append(line1);
            curves.Append(line2);
            curves.Append(line3);
            curves.Append(line4);

            // move by zOffset
            XYZ translation = new XYZ(0, 0, zOffset);
            Transform offsetTransform = Transform.CreateTranslation(translation);
            CurveLoop offsetCurveLoop = CurveLoop.CreateViaTransform(curves, offsetTransform);

            // extrude -z
            return GeometryCreationUtilities.CreateExtrusionGeometry(
                new List<CurveLoop> { offsetCurveLoop },
                XYZ.BasisZ.Negate(),
                floorThickness);
        }

        private Solid GetOrCreateRoomFloor(Document doc, Room room, Level level, out Floor outFloor, out double offset)
        {
            SpatialElementBoundaryOptions options = new SpatialElementBoundaryOptions();
            IList<IList<BoundarySegment>> boundaries = room.GetBoundarySegments(options);
            if (boundaries == null || boundaries.Count == 0)
            {
                outFloor = null;
                offset = 0;
                return null;
            }

            // boundary loops
            List<CurveLoop> loops = new List<CurveLoop>();

            foreach (IList<BoundarySegment> boundaryList in boundaries)
            {
                CurveLoop loop = new CurveLoop();
                foreach (BoundarySegment segment in boundaryList)
                {
                    loop.Append(segment.GetCurve());
                }
                loops.Add(loop);
            }

            // no existing floor
            double floorThickness = 0.15;
            offset = 0;

            // all floors
            FilteredElementCollector floors = new FilteredElementCollector(doc, doc.ActiveView.Id)
                .OfCategory(BuiltInCategory.OST_Floors)
                .OfClass(typeof(Floor));

            // intersected floor if exist
            XYZ roomCentroid = (room.Location as LocationPoint).Point;
            Line intersectionLine = Line.CreateBound(
                roomCentroid,
                new XYZ(roomCentroid.X, roomCentroid.Y, roomCentroid.Z - 10));

            foreach (Floor floor in floors)
            {
                GeometryElement floorGeom = floor.get_Geometry(new Options());
                if (floorGeom == null) continue;

                foreach (GeometryObject obj in floorGeom)
                {
                    Solid floorSolid = obj as Solid;
                    if (floorSolid != null && floorSolid.Volume > 0)
                    {
                        SolidCurveIntersection intersection = floorSolid.IntersectWithCurve(
                            intersectionLine,
                            new SolidCurveIntersectionOptions());

                        if (intersection != null && intersection.SegmentCount > 0)
                        {
                            outFloor = floor;
                            floorThickness = floor.get_Parameter(BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM).AsDouble();
                            offset = floor.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM).AsDouble();

                            if (offset != 0)
                            {
                                XYZ translation = new XYZ(0, 0, offset);
                                Transform offsetTransform = Transform.CreateTranslation(translation);

                                List<CurveLoop> offsetLoops = new List<CurveLoop>();
                                foreach (CurveLoop loop in loops)
                                {
                                    offsetLoops.Add(CurveLoop.CreateViaTransform(loop, offsetTransform));
                                }

                                return GeometryCreationUtilities.CreateExtrusionGeometry(
                                    offsetLoops,
                                    XYZ.BasisZ.Negate(),
                                    floorThickness);
                            }

                            return floorSolid;
                        }
                    }
                }
            }

            // no floor , new one
            Solid RoomFloorSolid = GeometryCreationUtilities.CreateExtrusionGeometry(
                loops,
                XYZ.BasisZ.Negate(),
                floorThickness);

            outFloor = null;
            offset = 0;
            return RoomFloorSolid;
        }

        private Floor CreateFloorFromSolid(Document doc, Solid solid, string floorTypeName, Level level)
        {
            FloorType floorType = new FilteredElementCollector(doc)
                                    .OfClass(typeof(FloorType))
                                    .Cast<FloorType>()
                                    .FirstOrDefault(f => f.Name == floorTypeName);

            if (floorType == null)
            {
                floorType = new FilteredElementCollector(doc)
                             .OfClass(typeof(FloorType))
                             .Cast<FloorType>()
                             .FirstOrDefault();

                if (floorType == null)
                {
                    TaskDialog.Show("Error", "no floor types found in the document.");
                    return null;
                }
            }

            Face topFace = null;

            // top face 
            foreach (Face face in solid.Faces)
            {
                XYZ normal = face.ComputeNormal(new UV(0.5, 0.5));
                if (normal.IsAlmostEqualTo(XYZ.BasisZ, 0.01))
                {
                    topFace = face;
                    break;
                }
            }

            if (topFace == null)
            {
                TaskDialog.Show("Error", "top face error.");
                return null;
            }

            IList<CurveLoop> loops = topFace.GetEdgesAsCurveLoops();
            if (loops == null || loops.Count == 0)
            {
                TaskDialog.Show("Error", "no curve loops found on top face.");
                return null;
            }

            try
            {
                return Floor.Create(doc, loops, floorType.Id, level.Id);
            }
            catch (Exception ex)
            {
                TaskDialog.Show("error creating floor", ex.Message);
                return null;
            }
        }

        private void CopyFloorParameters(Floor source, Floor target)
        {
            Parameter sourceParam = source.LookupParameter("Mark");
            Parameter targetParam = target.LookupParameter("Mark");

            if (sourceParam != null && targetParam != null && !targetParam.IsReadOnly && sourceParam.HasValue)
            {
                targetParam.Set(sourceParam.AsString());
            }
        }

        #endregion
    }
}