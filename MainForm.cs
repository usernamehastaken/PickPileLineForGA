using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;

namespace PickPileLineForGA
{
    public partial class MainForm : System.Windows.Forms.Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void splitContainer1_Panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            //this.Hide();
            ElementId elementId = myFuncs.UIDocument.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element).ElementId;
            Connector connectorFirstPick = myFuncs.connectorFirstPick();
            if (!(connectorFirstPick is null))
            {
                myFuncs.getPipeLineFromConnector(connectorFirstPick);
            }
        }
    }
}
