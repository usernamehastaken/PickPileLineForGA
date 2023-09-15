using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using System.Windows.Forms;

namespace PickPileLineForGA
{
    public class myFuncs
    {
        public static UIApplication uIApplication { get; set; }
        public static UIDocument UIDocument { get {return uIApplication.ActiveUIDocument;}}
        public static Document document { get { return UIDocument.Document; } }

        //选取第一个元素，得到唯一一个连接的连接件
        //不唯一报错
        public static void pickFirstElement()
        {

        }
        
        public static List<Connector> GetConnectors(Connector connector)
        {
            List<Connector> connectors = new List<Connector>();
            Element elementOwner = connector.Owner;
            MessageBox.Show(elementOwner.GetType().ToString());
            //switch (elementOwner.Category.get)
            //{
            //    default:
            //        break;
            //}
            return connectors;
        }

        private static bool Duct(Element elementOwner)
        {
            throw new NotImplementedException();
        }

        public static void firstPick()
        {
            Reference reference = UIDocument.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element);
            Element element = document.GetElement(reference.ElementId);

            BuiltInCategory builtInCategory = (BuiltInCategory)element.Category.Id.IntegerValue;
            switch (builtInCategory)
            {
                case BuiltInCategory.OST_DuctCurves:
                    Duct duct = (Duct)element;
                    break;
                case BuiltInCategory.OST_PipeCurves:
                    Pipe pipe = (Pipe)element;
                    break;
                default:
                    FamilyInstance familyInstance = (FamilyInstance)element;
                    break;
            }
            MessageBox.Show(builtInCategory.ToString());
        }
    }
}
