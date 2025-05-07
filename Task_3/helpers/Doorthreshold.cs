using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;


namespace Task_3.Helpers
{
    public class Doorthreshold
    {
        public XYZ Locatin { get; set; }

        public double Width { get; set; }

        public double Depth { get; set; }

        public Room Room { get; set; }

        public Wall HostWall { get; set; }

        public FamilyInstance Door { get; set; }

    }
}