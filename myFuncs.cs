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

        public static List<List<jieDian>> guanwang = new List<List<jieDian>>();
        public static List<List<jieDian>> linshiguanwang = new List<List<jieDian>>();
        //选取第一个元素，得到唯一一个连接的连接件
        //不唯一报错
        
        public static Connector connectorFirstPick()//选择管道的汇总处，选出未连接的连接件，并做管网及临时管网的初始化
        {
            Reference reference = UIDocument.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element);
            Element element = document.GetElement(reference.ElementId);
            Connector connectorReturn=null; //未连接的连接件
            //MessageBox.Show(element.Id.ToString());
            //throw new Exception();
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
                        return connectorReturn;
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
                        return connectorReturn;
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
                        return connectorReturn;
                    }
                    break;
            }

            guanwang.Clear();
            linshiguanwang.Clear();
            List<jieDian> zhilu = new List<jieDian>();
            jieDian jiedian = new jieDian();
            jiedian.owerId = element.Id;
            jiedian.connectorID = connectorReturn.Id;
            zhilu.Add(jiedian);
            linshiguanwang.Add(zhilu);//完成循环初始化需要的支路集合

            return connectorReturn;
        }

        public static void getPipeLineFromConnector()//**删除输入
        {
            //MessageBox.Show(guanwang[true].Count.ToString());
            //取guanwang最后一个元素进行循环，直至为空
            while (linshiguanwang.Count>0)
            {
                List<jieDian> zhilu = linshiguanwang[linshiguanwang.Count - 1];
                jieDian jiedian = zhilu[zhilu.Count - 1];//末端节点一定是refconnector的来
                //先取同一主体的不同节点
                List<jieDian> jieDians = getOwnerOtherConnectors(jiedian);
                if (jieDians.Count==0)//管道末端
                {
                    guanwang.Add(zhilu);
                    linshiguanwang.Remove(zhilu);
                }
                if (jieDians.Count>0)
                {
                    foreach (jieDian item in jieDians)
                    {
                        List<jieDian> xinzhilu = new List<jieDian>(zhilu.ToArray());
                        xinzhilu.Add(item);
                        linshiguanwang.Add(xinzhilu);
                    }
                    linshiguanwang.Remove(zhilu);

                    //后取不同主体的相关连接件
                    zhilu = linshiguanwang[linshiguanwang.Count - 1];
                    jiedian = zhilu[zhilu.Count - 1];//末端节点一定是refconnector的来
                                                     //先取同一主体的不同节点
                    jieDian xinjieDian = getRefConnector(jiedian);
                    if (xinjieDian == null)
                    {
                        guanwang.Add(zhilu);
                        linshiguanwang.Remove(zhilu);
                    }
                    if (xinjieDian != null)
                    {
                        List<jieDian> xinzhilu = new List<jieDian>(zhilu.ToArray());
                        xinzhilu.Add(xinjieDian);
                        linshiguanwang.Add(xinzhilu);
                        linshiguanwang.Remove(zhilu);
                    }
                }
                
            }
        }

        public static jieDian getRefConnector(jieDian jieDianIn)//输入一个连接件，返回与其相连接的不同主体相关连接件(唯一,可为空)
        {
            jieDian jieDianReturn = null;
            Element element = document.GetElement(jieDianIn.owerId);
            BuiltInCategory builtInCategory = (BuiltInCategory)element.Category.Id.IntegerValue;
            Connector connector = null;
            switch (builtInCategory)
            {
                case BuiltInCategory.OST_DuctCurves:
                    Duct duct = (Duct)element;
                    foreach (Connector item in duct.ConnectorManager.Connectors)
                    {
                        if (item.Id == jieDianIn.connectorID)
                        {
                            connector = item;
                        }
                    }
                    break;
                case BuiltInCategory.OST_PipeCurves:
                    Pipe pipe = (Pipe)element;
                    foreach (Connector item in pipe.ConnectorManager.Connectors)
                    {
                        if (item.Id == jieDianIn.connectorID)
                        {
                            connector = item;
                        }
                    }
                    break;
                default:
                    try
                    {
                        FamilyInstance familyInstance = (FamilyInstance)element;
                        foreach (Connector item in familyInstance.MEPModel.ConnectorManager.Connectors)
                        {
                            if (item.Id == jieDianIn.connectorID)
                            {
                                connector = item;
                            }
                        }
                    }
                    catch (Exception)
                    {

                        return jieDianReturn;
                    }
                    
                    break;
            }
            foreach (Connector item in connector.AllRefs)
            {
                if (item.Owner.Id != jieDianIn.owerId & item.ConnectorType==ConnectorType.End)
                {
                    jieDianReturn = new jieDian();
                    jieDianReturn.owerId = item.Owner.Id;
                    jieDianReturn.connectorID = item.Id;
                }
            }
            return jieDianReturn;
        }

        public static List<jieDian> getOwnerOtherConnectors(jieDian jieDianIn)//输入一个连接件，返回同一主体的其他相关连接件（不唯一,可为空）
        {
            List<jieDian> jieDiansReturn = new List<jieDian>();
            Element element = document.GetElement(jieDianIn.owerId);
            BuiltInCategory builtInCategory = (BuiltInCategory)element.Category.Id.IntegerValue;
            Connector connector = null;
            switch (builtInCategory)
            {
                case BuiltInCategory.OST_DuctCurves:
                    Duct duct = (Duct)element;
                    foreach (Connector item in duct.ConnectorManager.Connectors)
                    {
                        if (item.Id == jieDianIn.connectorID)
                        {
                            connector = item;
                        }
                    }
                    break;
                case BuiltInCategory.OST_PipeCurves:
                    Pipe pipe = (Pipe)element;
                    foreach (Connector item in pipe.ConnectorManager.Connectors)
                    {
                        if (item.Id == jieDianIn.connectorID)
                        {
                            connector = item;
                        }
                    }
                    break;
                default:
                    try
                    {
                        FamilyInstance familyInstance = (FamilyInstance)element;
                        foreach (Connector item in familyInstance.MEPModel.ConnectorManager.Connectors)
                        {
                            if (item.Id == jieDianIn.connectorID)
                            {
                                connector = item;
                            }
                        }
                    }
                    catch (Exception)
                    {

                        return jieDiansReturn;
                    }
                    break;
            }
            foreach (Connector item in connector.ConnectorManager.Connectors)
            {
                if (item.Id != jieDianIn.connectorID & item.ConnectorType==ConnectorType.End)
                {
                    jieDiansReturn.Add(new jieDian { owerId = element.Id, connectorID = item.Id });
                }
            }
            return jieDiansReturn;
        }
    }

    public class jieDian
    {
        public ElementId owerId { get; set; }
        public int connectorID { get; set; }
    }
    //public class zhiLu
    //{
    //    public List<jieDian> value { get; set; }
    //}
}
