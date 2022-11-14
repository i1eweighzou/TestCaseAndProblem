using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace TestCaseAndProblem
{
    public partial class ThisWorkbook
    {
        private ActionsPaneControlTestCase actionsPaneControlTestCase;

        internal ActionsPaneControlTestCase ActionsPaneControlTestCase { get => actionsPaneControlTestCase; set => actionsPaneControlTestCase = value; }

        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            ActionsPaneControlTestCase = new ActionsPaneControlTestCase();
            return Globals.Factory.GetRibbonFactory().CreateRibbonManager(
                    new Microsoft.Office.Tools.Ribbon.IRibbonExtension[] { new RibbonTestcaseAndProblem() });
        }
        private void ThisWorkbook_Startup(object sender, System.EventArgs e)
        {
            Globals.ChooseProblemSheet.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            Globals.EditItemsSheet.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            Globals.EditTestRecordSheet.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            Globals.EditProblemSheet.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            Globals.TestRecordAndProblemSheet.Visible = Excel.XlSheetVisibility.xlSheetHidden;
        }

        private void ThisWorkbook_Shutdown(object sender, System.EventArgs e)
        {

        }

        #region VSTO 设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisWorkbook_Startup);
            this.Shutdown += new System.EventHandler(ThisWorkbook_Shutdown);
        }

        #endregion

    }
}
