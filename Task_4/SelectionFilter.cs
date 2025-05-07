using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;

namespace Task4
{
    public class SelectionFilter : ISelectionFilter
    {
        private string categoryName;
        public SelectionFilter(string categoryName)
        {
            this.categoryName = categoryName;
        }
        public bool AllowElement(Element elem)
        {
            return elem.Category != null && elem.Category.Name == categoryName;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            return false;
        }
    }
}
