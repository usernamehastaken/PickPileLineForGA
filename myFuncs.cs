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

        public static Connector connectorFirstPick()//选择管道的汇总处，选出未连接的连接件，未作初始输入条件，开始管道拾取循环
        {
            Reference reference = UIDocument.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element);
            Element element = document.GetElement(reference.ElementId);
            Connector connectorReturn=null;

            BuiltInCategory builtInCategory = (BuiltInCategory)element.Category.Id.IntegerValue;
            switch (builtInCategory)
            {
                case BuiltInCategory.OST_DuctCurves:
                    Duct duct = (Duct)element;
                    if (duct.ConnectorManager.UnusedConnectors.Size==1)
                    {
                        foreach (Connector item in duct.ConnectorManager.UnusedConnectors)
                        {
                            connectorReturn = item;
                        }
                    }
                    else
                    {
                        MessageBox.Show("不满足管道系统汇总末端条件，未发现‘未使用的连接件’");
                    }
                    break;
                case BuiltInCategory.OST_PipeCurves:
                    Pipe pipe = (Pipe)element;
                    if (pipe.ConnectorManager.UnusedConnectors.Size == 1)
                    {
                        foreach (Connector item in pipe.ConnectorManager.UnusedConnectors)
                        {
                            connectorReturn = item;
                        }
                    }
                    else
                    {
                        MessageBox.Show("不满足管道系统汇总末端条件，未发现唯一‘未使用的连接件’");
                    }
                    break;
                default:
                    FamilyInstance familyInstance = (FamilyInstance)element;
                    if (familyInstance.MEPModel.ConnectorManager.UnusedConnectors.Size==1)
                    {
                        foreach (Connector item in familyInstance.MEPModel.ConnectorManager.UnusedConnectors)
                        {
                            connectorReturn = item;
                        }
                    }
                    else
                    {
                        MessageBox.Show("不满足管道系统汇总末端条件，未发现唯一‘未使用的连接件’");
                    }
                    break;
            }

            return connectorReturn;
        }

        public static void getPipeLineFromConnector(Connector connectorIn)
        {
            Element elementOwner = connectorIn.Owner;
            ConnectorSet connectorSet = connectorIn.ConnectorManager.Connectors;
            List<Connector> connectorsOther = new List<Connector>();
            foreach (Connector item in connectorSet)
            {
                if (item!=connectorIn)
                {
                    connectorsOther.Add(item);
                }
            }
            
        }
    }
}
