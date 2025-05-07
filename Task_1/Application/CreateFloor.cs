using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Task_1.HelperMethods;
using Task_1.Model;

namespace Task_1.Application
{
    [Transaction(TransactionMode.Manual)]
    public class CreateFloor : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;

            try
            {
                // Get a default FloorType
                FloorType floorType = new FilteredElementCollector(doc)
                    .OfClass(typeof(FloorType))
                    .WhereElementIsElementType()
                    .Cast<FloorType>()
                    .FirstOrDefault();

                if (floorType == null)
                {
                    message = "No floor type found in the document.";
                    return Result.Failed;
                }

                // Get the level from active view
                Level level = doc.ActiveView.GenLevel;
                if (level == null)
                {
                    message = "Active view has no associated level.";
                    return Result.Failed;
                }

                // Retrieve user-defined lines and generate a valid CurveLoop
                List<Line> boundaryLines = UserInputs.GetLines();
                IList<CurveLoop> curveLoops = Logic.ProcessCurveLoop(boundaryLines);

                if (curveLoops == null || curveLoops.Count == 0)
                {
                    TaskDialog.Show("Invalid Input", "Provided curves could not form a valid closed loop.");
                    return Result.Failed;
                }

                using (Transaction tx = new Transaction(doc, "Create Floor"))
                {
                    tx.Start();

                    // Convert first CurveLoop to CurveArray
                    CurveArray curveArray = new CurveArray();
                    foreach (Curve c in curveLoops.First())
                    {
                        curveArray.Append(c);
                    }

                    // Create the floor
                    Floor newFloor = doc.Create.NewFloor(curveArray, floorType, level, false);

                    tx.Commit();
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = $"An error occurred: {ex.Message}";
                return Result.Failed;
            }
        }
    }
}
