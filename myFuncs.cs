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
        public static UIDocument UIDocument { get { return uIApplication.ActiveUIDocument; } }
        public static Document document { get { return UIDocument.Document; } }

        public static List<List<jieDian>> guanwang = new List<List<jieDian>>();
        public static List<List<jieDian>> linshiguanwang = new List<List<jieDian>>();
        //选取第一个元素，得到唯一一个连接的连接件
        //不唯一报错

        public static Connector connectorFirstPick()//选择管道的汇总处，选出未连接的连接件，并做管网及临时管网的初始化
        {
            Reference reference = UIDocument.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element);
            Element element = document.GetElement(reference.ElementId);
            Connector connectorReturn = null; //未连接的连接件
            //MessageBox.Show(element.Id.ToString());
            //throw new Exception();
            BuiltInCategory builtInCategory = (BuiltInCategory)element.Category.Id.IntegerValue;
            switch (builtInCategory)
            {
                case BuiltInCategory.OST_DuctCurves:
                    Duct duct = (Duct)element;
                    if (duct.ConnectorManager.UnusedConnectors.Size == 1)
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
                    if (familyInstance.MEPModel.ConnectorManager.UnusedConnectors.Size == 1)
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

        public static void getPipeLineFromConnector()//**得到所有支路连接件
        {
            //MessageBox.Show(guanwang[true].Count.ToString());
            //取guanwang最后一个元素进行循环，直至为空
            while (linshiguanwang.Count > 0)
            {
                List<jieDian> zhilu = linshiguanwang[linshiguanwang.Count - 1];
                jieDian jiedian = zhilu[zhilu.Count - 1];//末端节点一定是refconnector的来
                //先取同一主体的不同节点
                List<jieDian> jieDians = getOwnerOtherConnectors(jiedian);
                if (jieDians.Count == 0)//管道末端
                {
                    guanwang.Add(zhilu);
                    linshiguanwang.Remove(zhilu);
                }
                if (jieDians.Count > 0)
                {
                    foreach (jieDian item in jieDians)
                    {
                        List<jieDian> xinzhilu = new List<jieDian>(zhilu.ToArray());//同一主体不同连接件加入管网后，应都循环找出对应连接件
                        xinzhilu.Add(item);
                        linshiguanwang.Add(xinzhilu);

                        jieDian xinjieDian = getRefConnector(item);
                        if (xinjieDian == null)
                        {
                            guanwang.Add(xinzhilu);
                            linshiguanwang.Remove(xinzhilu);
                        }
                        if (xinjieDian != null)
                        {
                            List<jieDian> gengxinzhilu = new List<jieDian>(xinzhilu.ToArray());
                            gengxinzhilu.Add(xinjieDian);
                            linshiguanwang.Add(gengxinzhilu);
                            linshiguanwang.Remove(xinzhilu);
                        }
                    }
                    linshiguanwang.Remove(zhilu);
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
                if (item.Owner.Id != jieDianIn.owerId & item.ConnectorType == ConnectorType.End)
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
            ConnectorSet connectorSet = null;
            switch (builtInCategory)
            {
                case BuiltInCategory.OST_DuctCurves:
                    Duct duct = (Duct)element;
                    connectorSet = duct.ConnectorManager.Connectors;
                    break;
                case BuiltInCategory.OST_PipeCurves:
                    Pipe pipe = (Pipe)element;
                    connectorSet = pipe.ConnectorManager.Connectors;
                    break;
                default:
                    try
                    {
                        FamilyInstance familyInstance = (FamilyInstance)element;
                        connectorSet = familyInstance.MEPModel.ConnectorManager.Connectors;
                    }
                    catch (Exception)
                    {

                        return jieDiansReturn;
                    }
                    break;
            }
            foreach (Connector item in connectorSet)
            {
                if (item.Id != jieDianIn.connectorID & item.ConnectorType != ConnectorType.Logical)
                {
                    jieDiansReturn.Add(new jieDian { owerId = element.Id, connectorID = item.Id });
                }
            }
            return jieDiansReturn;
        }

        public static void getListAllEndInfo(MainForm mainForm)
        {
            mainForm.dataGridView1.Rows.Clear();
            foreach (List<jieDian> item in guanwang)
            {
                jieDian jieDian = item[item.Count - 1];

                Element element = document.GetElement(jieDian.owerId);
                BuiltInCategory builtInCategory = (BuiltInCategory)element.Category.Id.IntegerValue;
                ConnectorSet connectorSet = null;
                Connector connector = null;
                switch (builtInCategory)
                {
                    case BuiltInCategory.OST_DuctCurves:
                        Duct duct = (Duct)element;
                        connectorSet = duct.ConnectorManager.Connectors;
                        break;
                    case BuiltInCategory.OST_PipeCurves:
                        Pipe pipe = (Pipe)element;
                        connectorSet = pipe.ConnectorManager.Connectors;
                        break;
                    default:
                        FamilyInstance familyInstance = (FamilyInstance)element;
                        connectorSet = familyInstance.MEPModel.ConnectorManager.Connectors;
                        break;
                }
                foreach (Connector citem in connectorSet)
                {
                    if (citem.Id == jieDian.connectorID)
                    {
                        connector = citem;
                    }
                }

                int rowindex = mainForm.dataGridView1.Rows.Add();
                mainForm.dataGridView1.Rows[rowindex].Cells["ID"].Value = element.Id;
                if (connector.Shape == ConnectorProfileType.Round)
                {
                    mainForm.dataGridView1.Rows[rowindex].Cells["Size"].Value = UnitUtils.Convert(connector.Radius*2, DisplayUnitType.DUT_DECIMAL_FEET, DisplayUnitType.DUT_MILLIMETERS);
                }
                if (connector.Shape==ConnectorProfileType.Rectangular)
                {
                    mainForm.dataGridView1.Rows[rowindex].Cells["Size"].Value = UnitUtils.Convert(connector.Width, DisplayUnitType.DUT_DECIMAL_FEET, DisplayUnitType.DUT_MILLIMETERS) + "x" + UnitUtils.Convert(connector.Height, DisplayUnitType.DUT_DECIMAL_FEET, DisplayUnitType.DUT_MILLIMETERS);
                }
                mainForm.dataGridView1.Rows[rowindex].Cells["Volume"].Value = get_Parameters("HC_AirFlow", element.Id);
            }
        }

        public static void getListAllzhiluInfo(MainForm mainForm)
        {
            //MessageBox.Show(mainForm.dataGridView1.Rows[1].Cells["Volume"].Value.ToString());
            for (int i = 0; i < mainForm.dataGridView1.Rows.Count; i++)
            {
                try
                {
                    string value = mainForm.dataGridView1.Rows[i].Cells["Volume"].Value.ToString().Trim();
                    if (value=="")
                    {
                        MessageBox.Show("有末端流量为0，请设置流量，再重新运行‘1.管网信息提取’");
                        return;
                    }
                }
                catch (Exception)
                {
                    MessageBox.Show("有末端流量为0，请设置流量，再重新运行‘1.管网信息提取’");
                    return;
                }
            }
            //检验结束
            mainForm.dataGridView1.Rows.Clear();
            int rowIndex;
            int num = 0;
            foreach (List<jieDian> zhilu in guanwang)
            {
                num = num + 1;
                for (int i = zhilu.Count-1; i >= 0; i--)
                {
                    jieDian jieDian = zhilu[i];//从末端算起
                    Element element = document.GetElement(jieDian.owerId);//得到图元
                    BuiltInCategory builtInCategory = (BuiltInCategory)element.Category.Id.IntegerValue;
                    Duct duct = null;
                    Pipe pipe = null;
                    FamilyInstance familyInstance = null;

                    #region 判断连接件主体及得到对应节点连接件
                    Connector connector = null;
                    switch (builtInCategory)
                    {
                        case BuiltInCategory.OST_DuctCurves:
                            duct = (Duct)element;
                            foreach (Connector item in duct.ConnectorManager.Connectors)
                            {
                                if (item.Id == jieDian.connectorID)
                                {
                                    connector = item;
                                }
                            }
                            break;
                        case BuiltInCategory.OST_PipeCurves:
                            pipe = (Pipe)element;
                            foreach (Connector item in pipe.ConnectorManager.Connectors)
                            {
                                if (item.Id == jieDian.connectorID)
                                {
                                    connector = item;
                                }
                            }
                            break;
                        default:
                            familyInstance = (FamilyInstance)element;
                            foreach (Connector item in familyInstance.MEPModel.ConnectorManager.Connectors)
                            {
                                if (item.Id == jieDian.connectorID)
                                {
                                    connector = item;
                                }
                            }
                            break;
                    }
                    #endregion

                    string size = get_ConnectorSize(jieDian.owerId, jieDian.connectorID);

                    #region 得到第一个（末端）连接件的流量
                    string volume = "0";
                    if (i == zhilu.Count - 1)
                    {
                        volume = get_Parameters("HC_AirFlow", element.Id);
                    }
                    #endregion

                    string angle = "0";
                    //支路名称：PipeLineName
                    //ID：      ID
                    //尺寸：    Size
                    //流量：    Volume
                    //类型：    Type
                    //角度：    Angle
                    rowIndex = mainForm.dataGridView1.Rows.Add();
                    mainForm.dataGridView1.Rows[rowIndex].Cells["PipeLineName"].Value = "支路" + i.ToString();
                    mainForm.dataGridView1.Rows[rowIndex].Cells["ID"].Value = element.Id;
                    if (i==0)//最后一个连接件
                    {

                    }
                    else
                    {
                        if (jieDian.owerId==zhilu[i-1].owerId)//同一主体连接件
                        {
                            if (familyInstance!=null)
                            {
                                switch (familyInstance.MEPModel.ConnectorManager.Connectors.Size)
                                {
                                    case 1:
                                        mainForm.dataGridView1.Rows[rowIndex].Cells["Size"].Value = size;
                                        mainForm.dataGridView1.Rows[rowIndex].Cells["Volume"].Value = volume;
                                        mainForm.dataGridView1.Rows[rowIndex].Cells["Type"].Value = "End";
                                        mainForm.dataGridView1.Rows[rowIndex].Cells["Angle"].Value = angle;
                                        break;
                                    case 2:
                                        angle=g
                                        break;
                                    case 3:
                                        break;
                                    case 4:
                                        break;
                                }
                            }
                            else //风管与水管一致
                            {
                                mainForm.dataGridView1.Rows[rowIndex].Cells["Size"].Value = size;
                                mainForm.dataGridView1.Rows[rowIndex].Cells["Volume"].Value = volume;
                                mainForm.dataGridView1.Rows[rowIndex].Cells["Type"].Value = duct.GetType().Name;
                                mainForm.dataGridView1.Rows[rowIndex].Cells["Angle"].Value = angle;
                            }
                        }
                    }
                }
            }
        }

        public static string get_Parameters(string par_name, ElementId id)
        {
            Element element = document.GetElement(id);
            ParameterSet parameterSet = element.Parameters;
            foreach (Parameter par in parameterSet)
            {
                if (par.Definition.Name == par_name)
                {
                    //char[] s = { ' ' };
                    return par.AsString();
                }
            }
            throw new Exception("Error: Can not get the parameter of " + par_name);
        }
        public static string get_ConnectorSize(ElementId elementId,int connectorID)
        {
            Element element = document.GetElement(elementId);//得到图元
            BuiltInCategory builtInCategory = (BuiltInCategory)element.Category.Id.IntegerValue;
            Duct duct = null;
            Pipe pipe = null;
            FamilyInstance familyInstance = null;
            #region 判断连接件主体及得到对应节点连接件
            Connector connector = null;
            switch (builtInCategory)
            {
                case BuiltInCategory.OST_DuctCurves:
                    duct = (Duct)element;
                    foreach (Connector item in duct.ConnectorManager.Connectors)
                    {
                        if (item.Id == connectorID)
                        {
                            connector = item;
                        }
                    }
                    break;
                case BuiltInCategory.OST_PipeCurves:
                    pipe = (Pipe)element;
                    foreach (Connector item in pipe.ConnectorManager.Connectors)
                    {
                        if (item.Id == connectorID)
                        {
                            connector = item;
                        }
                    }
                    break;
                default:
                    familyInstance = (FamilyInstance)element;
                    foreach (Connector item in familyInstance.MEPModel.ConnectorManager.Connectors)
                    {
                        if (item.Id == connectorID)
                        {
                            connector = item;
                        }
                    }
                    break;
            }
            #endregion

            #region 得到连接件的Size
            string size = "";//连接件尺寸
            switch (connector.Shape)
            {
                case ConnectorProfileType.Invalid:
                    throw new Exception("管道尺寸不在软件考虑范围");
                case ConnectorProfileType.Round:
                    size = UnitUtils.Convert(connector.Radius, DisplayUnitType.DUT_DECIMAL_FEET, DisplayUnitType.DUT_MILLIMETERS).ToString();
                    break;
                case ConnectorProfileType.Rectangular:
                    size = UnitUtils.Convert(connector.Width, DisplayUnitType.DUT_DECIMAL_FEET, DisplayUnitType.DUT_MILLIMETERS).ToString() +
                        "x" + UnitUtils.Convert(connector.Height, DisplayUnitType.DUT_DECIMAL_FEET, DisplayUnitType.DUT_MILLIMETERS).ToString();
                    break;
                case ConnectorProfileType.Oval:
                    throw new Exception("椭圆管道尺寸不在软件考虑范围");
            }
            #endregion
            return size;
        }

        public static string get_ConnectorsAngle(ElementId elementId1,int connectorID1,ElementId elementId2,int connectorID2)
        {
            string angel = "0";
            FamilyInstance f1 = (FamilyInstance)document.GetElement(elementId1);
            FamilyInstance f2 = (FamilyInstance)document.GetElement(elementId2);
            Connector c1 = null;
            Connector c2 = null;
            foreach (Connector item in f1.MEPModel.ConnectorManager.Connectors)
            {
                if (item.Id == connectorID1)
                {
                    c1 = item;
                    break;
                }
            }
            foreach (Connector item in f2.MEPModel.ConnectorManager.Connectors)
            {
                if (item.Id == connectorID2)
                {
                    c2 = item;
                    break;
                }
            }
            
        }
        public static double get_Angle(XYZ p1, XYZ p2)
        {
            double angle = Math.Round(p1.AngleTo(p2) / Math.PI * 180, 2);
            if (angle > 90)
            {
                return 180 - angle;
            }
            return angle;
        }
        public class jieDian
        {
            public ElementId owerId { get; set; }
            public int connectorID { get; set; }
        }
    }
}
