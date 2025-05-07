using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace Task_1.Model
{
    internal class UserInputs
    {
        public static List<Line> GetLines()
        {
            return new List<Line>
            {
                //List of lines as in pdf with the same order
                Line.CreateBound(new XYZ(0,0,0), new XYZ(79,0,0)),
                Line.CreateBound(new XYZ(44,25,0), new XYZ(13,25,0)),
                Line.CreateBound(new XYZ(13,40,0), new XYZ(-8,40,0)),
                Line.CreateBound(new XYZ(55,34,0), new XYZ(55,10,0)),
                Line.CreateBound(new XYZ(79,34,0), new XYZ(55,34,0)),
                Line.CreateBound(new XYZ(0,20,0), new XYZ(0,0,0)),
                Line.CreateBound(new XYZ(55,10,0), new XYZ(44,12,0)),
                Line.CreateBound(new XYZ(-8,40,0), new XYZ(-8,20,0)),
                Line.CreateBound(new XYZ(79,0,0), new XYZ(79,34,0)),
                Line.CreateBound(new XYZ(44,12,0), new XYZ(44,25,0)),
                Line.CreateBound(new XYZ(-8,20,0), new XYZ(0,20,0)),
                Line.CreateBound(new XYZ(13,25,0), new XYZ(13,40,0)),
            };
        }

    }
}