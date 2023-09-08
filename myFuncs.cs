using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
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
        
        public static ElementId getAnotherElementID(ElementId elementId,Connector connector)
        {
            Element element = document.GetElement(elementId);//得到输入的ID对应的元素
            element
        }

        public static void tmp()
        {
            ;
        }
    }
}
