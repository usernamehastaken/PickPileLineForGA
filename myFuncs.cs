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
using System.IO;

namespace PickPileLineForGA
{
    //[Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class myFuncs
    {
        //此项为非模态窗体修改图元参数使用
        public static MainForm MyForm = null;
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
            jieDian jiedian = new jieDian
            {
                owerId = element.Id,
                connectorID = connectorReturn.Id
            };
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
            //判断是否有三通或四通直接作为支路末端
            List<ElementId> tmpL = new List<ElementId>();
            foreach (List<jieDian> item in guanwang)
            {
                tmpL.Add(item[item.Count - 1].owerId);
            }
            if (tmpL.Distinct().ToList().Count!=guanwang.Count)
            {
                MessageBox.Show("有三通或四通直接作为管网末端，请加末端或管段作为末端并重新运行");
                guanwang.Clear();
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
            //判断是否已经运行功能1
            if (mainForm.dataGridView1.Rows.Count<1)
            {
                MessageBox.Show("表中信息为空，请先运行功能1");
                return;
            }
            //检测是否有末端未设置流量
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
            //开始进入实际主程序
            mainForm.dataGridView1.Rows.Clear();
            int num = 0;//用于支路名称
            foreach (List<jieDian> zhilu in guanwang)//从guanwang里的zhilu开始循环
            {
                num = num + 1;
                for (int i = zhilu.Count-1; i >= 0; i--)//从支路的末端jiedian算起
                {
                    #region 循环主题的变量定义
                    jieDian jieDian = zhilu[i];//支路节点
                    Element element = document.GetElement(jieDian.owerId);//得到图元
                    BuiltInCategory builtInCategory = (BuiltInCategory)element.Category.Id.IntegerValue;
                    Duct duct = null;
                    Pipe pipe = null;
                    FamilyInstance familyInstance = null;
                    #endregion
                    #region 得到当前节点的主体及对应连接件，判断连接件主体（后续用isnull判断）
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
                    //得到连接件的尺寸
                    string size = get_Connector_Size(jieDian.owerId, jieDian.connectorID);
                    string volume = get_Parameters("HC_AirFlow", element.Id);
                    string angle = "0";
                    string length = "0";
                    try
                    {
                        length = element.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsValueString();
                    }
                    catch (Exception)
                    {
                        ;
                    }
                    //末端是三通及四通的情况已经排除
                    //判断是否为最后的总管
                    //判断是否与下一个节点同图元，如果是直接进行下一个循环直到不是
                    if (i==zhilu.Count-1)//最后一个可以往后判断但是无法往前判断
                    {
                        if (jieDian.owerId != zhilu[i-1].owerId)//末端唯一节点
                        {
                            set_mainForm_data(mainForm, "支路" + num, element.Id, size, length, volume, "End", angle);
                        }
                    }
                    else if (i>0)//非末端及最后节点，或末端非唯一连接件
                    {
                        if (jieDian.owerId != zhilu[i-1].owerId)//图元的最后一个连接件进行图元计算,如果一样直接跳过
                        {
                            if (familyInstance!=null)
                            {
                                switch (familyInstance.MEPModel.ConnectorManager.Connectors.Size)
                                {
                                    case 2://弯头，变径,计算角度/阀门支架附件等均归于Reducer
                                        angle = get_ConnectorsAngle(jieDian.owerId, jieDian.connectorID, zhilu[i + 1].owerId, zhilu[i + 1].connectorID);
                                        if (angle=="0")
                                        {
                                            set_mainForm_data(mainForm, "支路" + num, element.Id, size, length, volume, "Reducer", angle);
                                        }
                                        else
                                        {
                                            set_mainForm_data(mainForm, "支路" + num, element.Id, size, length, volume, "Elbow", angle);
                                        }
                                        break;
                                    case 3://三通，计算角度
                                        angle = get_ConnectorsAngle(jieDian.owerId, jieDian.connectorID, zhilu[i + 1].owerId, zhilu[i + 1].connectorID);
                                        set_mainForm_data(mainForm, "支路" + num, element.Id, size, length, volume, "Tee", angle);
                                        break;
                                    case 4://四通，计算角度
                                        angle = get_ConnectorsAngle(jieDian.owerId, jieDian.connectorID, zhilu[i + 1].owerId, zhilu[i + 1].connectorID);
                                        set_mainForm_data(mainForm, "支路" + num, element.Id, size, length, volume, "Cross", angle);
                                        break;
                                } ;
                            }
                            else//只能是水管或风管
                            {
                                if (duct != null)
                                {
                                    set_mainForm_data(mainForm, "支路" + num, element.Id, size,length, volume, "Duct", angle);
                                }
                                else
                                {
                                    set_mainForm_data(mainForm, "支路" + num, element.Id, size,length, volume, "Pipe", angle);
                                }
                            }
                        }
                    }
                    if (i==0)
                    {
                        if (jieDian.owerId != zhilu[i+1].owerId)//末端只有一个连接件，非弯头，风管，三通，四通，只能是族
                        {
                            set_mainForm_data(mainForm, "支路" + num, element.Id, size,length, volume, "总管", angle);
                        }
                        else
                        {
                            if (familyInstance != null)
                            {
                                switch (familyInstance.MEPModel.ConnectorManager.Connectors.Size)
                                {
                                    case 2://弯头，变径,计算角度/阀门支架附件等均归于Reducer
                                        angle = get_ConnectorsAngle(jieDian.owerId, jieDian.connectorID, zhilu[i + 1].owerId, zhilu[i + 1].connectorID);
                                        if (angle == "0")
                                        {
                                            set_mainForm_data(mainForm, "支路" + num, element.Id, size, length, volume, "Reducer", angle);
                                        }
                                        else
                                        {
                                            set_mainForm_data(mainForm, "支路" + num, element.Id, size, length, volume, "Elbow", angle);
                                        }
                                        break;
                                    case 3://三通，计算角度
                                        angle = get_ConnectorsAngle(jieDian.owerId, jieDian.connectorID, zhilu[i + 1].owerId, zhilu[i + 1].connectorID);
                                        set_mainForm_data(mainForm, "支路" + num, element.Id, size,length, volume, "Tee", angle);
                                        break;
                                    case 4://四通，计算角度
                                        angle = get_ConnectorsAngle(jieDian.owerId, jieDian.connectorID, zhilu[i + 1].owerId, zhilu[i + 1].connectorID);
                                        set_mainForm_data(mainForm, "支路" + num, element.Id, size,length, volume, "Cross", angle);
                                        break;
                                };
                            }
                            else//只能是水管或风管
                            {
                                if (duct != null)
                                {
                                    set_mainForm_data(mainForm, "支路" + num, element.Id, size, length, volume, "Duct", angle);
                                }
                                else
                                {
                                    set_mainForm_data(mainForm, "支路" + num, element.Id, size,length, volume, "Pipe", angle);
                                }
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
        public static string get_Connector_Size(ElementId elementId,int connectorID)
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
                    size = UnitUtils.Convert(connector.Radius*2, DisplayUnitType.DUT_DECIMAL_FEET, DisplayUnitType.DUT_MILLIMETERS).ToString();
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
        public static void set_mainForm_data(MainForm mainForm,string PipeLineName,ElementId ID,string Size,string Length,string Volume,string Type,string Angle)
        {
            //支路名称：PipeLineName
            //ID：      ID
            //尺寸：    Size
            //流量：    Volume
            //类型：    Type
            //角度：    Angle
            int rowIndex = mainForm.dataGridView1.Rows.Add();
            mainForm.dataGridView1.Rows[rowIndex].Cells["PipeLineName"].Value = PipeLineName;
            mainForm.dataGridView1.Rows[rowIndex].Cells["ID"].Value = ID;
            mainForm.dataGridView1.Rows[rowIndex].Cells["Volume"].Value = Volume;
            mainForm.dataGridView1.Rows[rowIndex].Cells["Type"].Value = Type;
            mainForm.dataGridView1.Rows[rowIndex].Cells["Angle"].Value = Angle;
            mainForm.dataGridView1.Rows[rowIndex].Cells["Size"].Value = Size;
            mainForm.dataGridView1.Rows[rowIndex].Cells["Length"].Value = Length;
        }
        public static string get_ConnectorsAngle(ElementId elementId1,int connectorID1,ElementId elementId2,int connectorID2)
        {
            FamilyInstance F1 = (FamilyInstance)document.GetElement(elementId1);
            Connector connector1 = null;
            FamilyInstance F2 = (FamilyInstance)document.GetElement(elementId2);
            Connector connector2 = null;
            foreach (Connector item in F1.MEPModel.ConnectorManager.Connectors)
            {
                if (item.Id==connectorID1)
                {
                    connector1 = item;
                    break;
                }
            }
            foreach (Connector item in F2.MEPModel.ConnectorManager.Connectors)
            {
                if (item.Id==connectorID2)
                {
                    connector2 = item;
                }
            }
            XYZ origin1 = F1.FacingOrientation;
            XYZ origin2 = F2.FacingOrientation;
            XYZ v1 = get_VectorFromConnector(connector1, origin1);
            XYZ v2 = get_VectorFromConnector(connector2, origin2);
            return get_Angle(v1, v2).ToString();
        }
        public static XYZ get_VectorFromConnector(Connector connector, XYZ origin)
        {
            double X = connector.CoordinateSystem.BasisX.X * origin.X + connector.CoordinateSystem.BasisX.Y * origin.Y + connector.CoordinateSystem.BasisX.Z * origin.Z;
            double Y = connector.CoordinateSystem.BasisY.X * origin.X + connector.CoordinateSystem.BasisY.Y * origin.Y + connector.CoordinateSystem.BasisY.Z * origin.Z;
            double Z = connector.CoordinateSystem.BasisZ.X * origin.X + connector.CoordinateSystem.BasisZ.Y * origin.Y + connector.CoordinateSystem.BasisZ.Z * origin.Z;
            return new XYZ(X, Y, Z);
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
        public static void set_flow_to_all_zhilu(MainForm mainForm)
        {
            int rowCount = mainForm.dataGridView1.Rows.Count;//总行数进行循环使用
            #region 对所有非末端管段进行流量赋值
            //末端集合用于循环使用
            List<jieDian> list_moduan = new List<jieDian>();
            foreach (List<jieDian> zhilu in guanwang)
            {
                list_moduan.Add(zhilu[zhilu.Count - 1]);
            }
            //将所有非末端流量进行赋值
            for (int i = 0; i < rowCount; i++)
            {
                if (i>0)
                {
                    if (mainForm.dataGridView1.Rows[i].Cells["PipeLineName"].Value.ToString() == mainForm.dataGridView1.Rows[i - 1].Cells["PipeLineName"].Value.ToString())
                    {
                        mainForm.dataGridView1.Rows[i].Cells["Volume"].Value = mainForm.dataGridView1.Rows[i - 1].Cells["Volume"].Value;
                    }
                }
            }
            #endregion

            #region 将相同ID的元素流量合并
            Dictionary<string, float> dic_flow = new Dictionary<string, float>();
            for (int i = 0; i < rowCount; i++)
            {
                if (dic_flow.Keys.Contains(mainForm.dataGridView1.Rows[i].Cells["ID"].Value.ToString())==true)
                {
                    dic_flow[mainForm.dataGridView1.Rows[i].Cells["ID"].Value.ToString()] += float.Parse(mainForm.dataGridView1.Rows[i].Cells["Volume"].Value.ToString());
                }
                else
                {
                    dic_flow[mainForm.dataGridView1.Rows[i].Cells["ID"].Value.ToString()] = float.Parse(mainForm.dataGridView1.Rows[i].Cells["Volume"].Value.ToString());
                }
            }
            //对表中单元格进行赋值
            for (int i = 0; i < rowCount; i++)
            {
                mainForm.dataGridView1.Rows[i].Cells["Volume"].Value = dic_flow[mainForm.dataGridView1.Rows[i].Cells["ID"].Value.ToString()].ToString();
                
            }
            #endregion
        }
        public static void data_to_csv(MainForm mainForm,string fullname)
        {
            //支路名称：PipeLineName
            //ID：      ID
            //尺寸：    Size
            //长度：    Length
            //流量：    Volume
            //类型：    Type
            //角度：    Angle
            if (File.Exists(fullname))
            {
                File.Delete(fullname);
            }
            using (FileStream fileStream = new FileStream(fullname, FileMode.OpenOrCreate, FileAccess.Write))
            {
                StreamWriter streamWriter=new StreamWriter(fileStream);
                int rowCount = mainForm.dataGridView1.RowCount;
                streamWriter.WriteLine("PipeLineName,ID,Size,Length,Volume,Type,Angle,");
                for (int i = 0; i < rowCount; i++)
                {
                    int columnCount = mainForm.dataGridView1.ColumnCount;
                    string strData = "";
                    for (int j = 0; j < columnCount; j++)
                    {
                        strData += mainForm.dataGridView1.Rows[i].Cells[j].Value.ToString()+",";
                    }
                    streamWriter.WriteLine(strData);
                }
                streamWriter.Close();
            }
            MessageBox.Show("文件已经导出，请登入平台进行计算！");
        }
        public static void csv_to_data(MainForm mainForm,string filename)
        {
            //支路名称：PipeLineName
            //ID：      ID
            //尺寸：    Size
            //长度:     Length
            //流量：    Volume
            //类型：    Type
            //角度：    Angle
            using (FileStream fileStream =new FileStream (filename,FileMode.Open,FileAccess.Read))
            {
                StreamReader reader=new StreamReader(fileStream);
                if (mainForm.dataGridView1.Rows.Count > 0)
                {
                    mainForm.dataGridView1.Rows.Clear();
                }
                string[] strData = reader.ReadToEnd().Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
                for (int i = 1; i < strData.Length; i++)
                {
                    string[] strValue = strData[i].Split(',');
                    int rowIndex = mainForm.dataGridView1.Rows.Add();
                    for (int j = 0; j < 7; j++)
                    {
                        mainForm.dataGridView1.Rows[rowIndex].Cells[j].Value = strValue[j];
                    }
                }
                reader.Close();
            }
        }
        public static void setValue_to_revitFile(UIApplication uIApplication)
        {
            MainForm mainForm = MyForm;
            using (Transaction transaction=new Transaction(document))
            {
                transaction.Start("PickPileLineForGA_setValue");
                try
                {
                    for (int i = 0; i < mainForm.dataGridView1.RowCount; i++)
                    {
                        Element element = document.GetElement(new ElementId(int.Parse(mainForm.dataGridView1.Rows[i].Cells["ID"].Value.ToString())));
                        List<Parameter> parameters = element.GetParameters("HC_AirFlow").ToList();
                        parameters[0].Set(mainForm.dataGridView1.Rows[i].Cells["Volume"].Value.ToString());
                    }
                    transaction.Commit();
                }
                catch (Exception)
                {
                    transaction.RollBack();
                    throw;
                }
            }
        }
        public class jieDian
        {
            //zhilu =list<jieDian>
            //guanwang=list<zhilu>
            public ElementId owerId { get; set; }
            public int connectorID { get; set; }
        }
    }
    public class ExecuteEventHandler : IExternalEventHandler
    {
        public string Name { get; set; }
        public Action<UIApplication> ExecuteAction { get; set; }
        public ExecuteEventHandler(string name)
        {
            Name = name;
        }
        public void Execute(UIApplication uIApp)
        {
            if (ExecuteAction != null)
            {
                //try
                //{
                ExecuteAction(uIApp);
                //}
                //catch (Exception e)
                //{
                //    //MessageBox.Show(e.Message);
                //}
            }
        }
        public string GetName()
        {
            return Name;
        }

    }
    public class test
    {
        public static void run()
        {
            Element element = myFuncs.document.GetElement(new ElementId(3744329));
            MessageBox.Show(element.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsValueString());
        }
    }
}
