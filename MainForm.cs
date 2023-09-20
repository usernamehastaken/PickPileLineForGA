using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace PickPileLineForGA
{
    public partial class MainForm : System.Windows.Forms.Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void DataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int rowindex = e.RowIndex;
            //MessageBox.Show(dataGridView1.Rows[rowindex].Cells["ID"].Value.ToString());
            ElementId elementId = new ElementId(int.Parse(this.dataGridView1.Rows[rowindex].Cells["ID"].Value.ToString()));
            myFuncs.UIDocument.Selection.SetElementIds(new List<ElementId> { elementId });
        }


        private void button1_Click(object sender, EventArgs e)
        {
            //ElementId elementId = myFuncs.UIDocument.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element).ElementId;
            Connector connectorFirstPick = myFuncs.connectorFirstPick();

            if (!(connectorFirstPick is null))
            {
                myFuncs.getPipeLineFromConnector();
            }
            MessageBox.Show("一共找到：" + myFuncs.guanwang.Count.ToString() + "支路");
            MessageBox.Show("列表为左右末端信息，请更改或核对流量数据!");
            myFuncs.getListAllEndInfo(this);
        }


        private void button2_Click(object sender, EventArgs e)
        {
            myFuncs.getListAllzhiluInfo(this);
        }
    }
}
