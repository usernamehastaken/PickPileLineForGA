using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace PickPileLineForGA
{
    
    public partial class MainForm : System.Windows.Forms.Form
    {
        public ExecuteEventHandler _executeEventHandler = null;
        public ExternalEvent _externalEvent = null;
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
            Connector connectorFirstPick = myFuncs.connectorFirstPick();

            if (!(connectorFirstPick is null))
            {
                myFuncs.getPipeLineFromConnector();
            }
            if (myFuncs.guanwang.Count > 0)
            {
                MessageBox.Show("一共找到：" + myFuncs.guanwang.Count.ToString() + "支路");
                MessageBox.Show("列表为左右末端信息，请更改或核对流量数据!");
                myFuncs.getListAllEndInfo(this);
            }
        }


        private void button2_Click(object sender, EventArgs e)
        {
            myFuncs.getListAllzhiluInfo(this);
            MessageBox.Show("进行各管段流量计算");
            myFuncs.set_flow_to_all_zhilu(this);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "|*.csv";
            saveFileDialog.Title = "导出数据";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                myFuncs.data_to_csv(this,saveFileDialog.FileName);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            myFuncs.MyForm = this;
            _executeEventHandler.ExecuteAction = null;
            _executeEventHandler.ExecuteAction += myFuncs.setValue_to_revitFile;
            _externalEvent.Raise();
            MessageBox.Show("修改完毕！");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "|*.csv";
            openFileDialog.Title = "导入数据";
            if (openFileDialog.ShowDialog()==DialogResult.OK)
            {
                myFuncs.csv_to_data(this, openFileDialog.FileName);
            }
        }

        private void process1_Exited(object sender, EventArgs e)
        {
                    }
    }
}
