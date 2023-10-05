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
            //test.run();
            //========================================
            MainForm mainForm = new MainForm();
            ExecuteEventHandler executeEventHandler = new ExecuteEventHandler("my exe");
            ExternalEvent externalEvent = ExternalEvent.Create(executeEventHandler);
            mainForm._executeEventHandler = executeEventHandler;
            mainForm._externalEvent = externalEvent;
            mainForm.Show();
            //========================================
            return Result.Succeeded;
        }
    }
}
