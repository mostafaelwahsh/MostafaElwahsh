using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Schema;
using  Task_2.Helpers;
using Task2.Utilities;

namespace Task2.Application
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        public static UIDocument CommandUIdoc { get; set; }
        public static Document CommandDoc { get; set; }
        public static Autodesk.Revit.ApplicationServices.Application CommandApp { get; set; }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            CommandApp = commandData.Application.Application;
            CommandUIdoc = commandData.Application.ActiveUIDocument;
            CommandDoc = CommandUIdoc.Document;

            IList<Autodesk.Revit.DB.BoundarySegment> boundarySegmentList = new List<BoundarySegment>();
            XYZ roomMidPoint = XYZ.Zero;

            try
            {
                var wallReference = CommandUIdoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element);
                var wallElement = CommandDoc.GetElement(wallReference) as Wall;

                var wcFamilySymbol = new FilteredElementCollector(CommandDoc)
                                     .OfCategory(BuiltInCategory.OST_GenericModel)
                                     .WhereElementIsElementType()
                                     .Where(a => a.Name == "ADA")
                                     .Cast<FamilySymbol>()
                                     .FirstOrDefault();


                var roomFound = RoomHelper.GetRoomsNextToSelectedWall(wallElement, out boundarySegmentList);

                var selectedWallCurveInRoom = GeometryHelper.GetWallCurveInsideRoom(boundarySegmentList, wallElement);

                var doorsLocPoints = RoomHelper.GetDoorsLocationInRoom(roomFound, boundarySegmentList);

                var familyLocationPoint = GeometryHelper.FarthestPointToCurve(doorsLocPoints, selectedWallCurveInRoom);

                var familyOrientation = GeometryHelper.GetCorrectFamilyOrientation(roomFound, familyLocationPoint, selectedWallCurveInRoom, out roomMidPoint);

                using (Transaction tr = new Transaction(CommandDoc, "Place Family"))
                {
                    tr.Start();

                    if (!wcFamilySymbol.IsActive)
                    {
                        wcFamilySymbol.Activate();
                    }

                    var familyCreated = CommandDoc.Create.NewFamilyInstance(familyLocationPoint, wcFamilySymbol, familyOrientation, wallElement, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                    FamilyInstanceHelper.FamilyOrientationHandler(CommandDoc, familyCreated, familyOrientation, selectedWallCurveInRoom, familyLocationPoint, roomMidPoint);


                    tr.Commit();
                }
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", ex.Message);
                return Result.Failed;

            }
        }




    }
}
