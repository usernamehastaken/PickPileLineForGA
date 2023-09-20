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
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand     //===========程序入口
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            myFuncs.uIApplication = commandData.Application;
            //========================================
            MainForm mainForm = new MainForm();
            mainForm.Show();
            //========================================
            return Result.Succeeded;
        }
    }
}
