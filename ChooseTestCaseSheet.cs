using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace TestCaseAndProblem
{
    public partial class ChooseTestCaseSheet
    {
        public Range orig_range;
        public string problem_id;
        private void ChooseTestCaseSheet_Startup(object sender, System.EventArgs e)
        {
        }

        private void ChooseTestCaseSheet_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO 设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.buttonTestCaseSubmit.Click += new System.EventHandler(this.buttonTestCaseSubmit_Click);
            this.checkBoxChoose.CheckedChanged += new System.EventHandler(this.checkBoxChoose_CheckedChanged);
            this.SelectionChange += new Microsoft.Office.Interop.Excel.DocEvents_SelectionChangeEventHandler(this.ChooseTestCaseSheet_SelectionChange);
            this.Startup += new System.EventHandler(this.ChooseTestCaseSheet_Startup);
            this.Shutdown += new System.EventHandler(this.ChooseTestCaseSheet_Shutdown);

        }

        #endregion

        private void ChooseTestCaseSheet_SelectionChange(Range Target)
        {
            int max_row = RangeUtils.get_max_row(Cells);
            if (Target.Row > max_row || Target.Row == 1)
            {
                return;
            }
            if (Target.Column == 3)
            {
                checkBoxChoose.Top = Target.Top;
                checkBoxChoose.Left = Target.Left;
                checkBoxChoose.Height = Target.Height;
                checkBoxChoose.Width = Target.Width;
                checkBoxChoose.Visible = true;
                checkBoxChoose.Tag = Target;
                if (String.Equals(Target.Text, "是"))
                {
                    checkBoxChoose.Checked = true;
                }
                else
                {
                    checkBoxChoose.Checked = false;
                }
            }
            else
            {
                checkBoxChoose.Visible = false;
            }
            buttonTestCaseSubmit.Top = Cells[Target.Row, 4].Top;
            buttonTestCaseSubmit.Left = Cells[Target.Row, 4].Left;
            buttonTestCaseSubmit.Visible = true;
        }

        private void buttonTestCaseSubmit_Click(object sender, EventArgs e)
        {
            Visible = XlSheetVisibility.xlSheetHidden;
            Globals.ProblemSheet.Visible = XlSheetVisibility.xlSheetVisible;
            Globals.ProblemSheet.Activate();
        }

        private void checkBoxChoose_CheckedChanged(object sender, EventArgs e)
        {
            int testcase_id_column = RangeUtils.get_column_by_title(Cells, "用例标识");
            Range range = checkBoxChoose.Tag as Range;
            String testcase_id = Cells[range.Row, testcase_id_column].Text;
            if (checkBoxChoose.Checked)
            {
                range.Value = "是";
                Globals.TestCaseAndProblemSheet.set_testcase_and_problem_cell(testcase_id, problem_id, "是");
            }
            else
            {
                range.Value = "";
                Globals.TestCaseAndProblemSheet.set_testcase_and_problem_cell(testcase_id, problem_id, "");
            }

            orig_range.Value = Globals.TestCaseAndProblemSheet.get_testcases_by_problem_id(problem_id);
        }
    }
}
