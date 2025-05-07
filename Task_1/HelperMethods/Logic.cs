using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

namespace Task_1.HelperMethods
{
    public class Logic
    {
        public static IList<CurveLoop> ProcessCurveLoop(List<Line> linesGiven)
        {
            IList<CurveLoop> curveLoopList = new List<CurveLoop>();
            // Attempt to construct a CurveLoop using the initial method
            CurveLoop curveLoop = TryConstructCurveLoop(linesGiven, CanConstructCurveLoop);

            if (curveLoop != null)
            {
                // The initial method succeeded
                curveLoopList.Add(curveLoop);
                return curveLoopList;
            }
            else
            {
                // The initial method failed, try an alternative method
                curveLoop = TryConstructCurveLoop(linesGiven, TryAlternativeMethod);

                if (curveLoop != null)
                {
                    // The alternative method succeeded
                    curveLoopList.Add(curveLoop);
                    return curveLoopList;
                }
                else
                {
                    // Both methods failed, return null or handle the failure
                    return null;
                }
            }
        }

        private static CurveLoop TryConstructCurveLoop(List<Line> linesGiven, Func<List<Line>, CurveLoop> constructionMethod)
        {
            try
            {
                // Attempt to construct a CurveLoop using the specified method
                return constructionMethod(linesGiven);
            }
            catch
            {
                // If an exception occurs, return null silently
                return null;
            }
        }

        private static CurveLoop CanConstructCurveLoop(List<Line> linesGiven)
        {

            CurveLoop curveLoop = new CurveLoop();
            foreach (var line in linesGiven)
            {
                curveLoop.Append(line);
            }
            return curveLoop;
        }

        public static CurveLoop TryAlternativeMethod(List<Line> lines)
        {
            // Sort lines so they connect end to start

            CurveLoop curveLoop = new CurveLoop();



            List<Line> chainageList = new List<Line> { lines.First() };

            for (int i = 0; i < lines.Count - 1; i++)
            {
                XYZ currentEndPoint = chainageList.Last().GetEndPoint(1);

                // Find the next line whose start point is equal to the current endpoint
                Line nextLine = lines.FirstOrDefault(line =>
                line.GetEndPoint(0).IsAlmostEqualTo(currentEndPoint));

                chainageList.Add(nextLine);

            }

            foreach (var line in chainageList)
            {
                XYZ endPoint1 = line.GetEndPoint(0);
                XYZ endPoint2 = line.GetEndPoint(1);
                curveLoop.Append(line);
            }

            return curveLoop;
        }

    }
}
