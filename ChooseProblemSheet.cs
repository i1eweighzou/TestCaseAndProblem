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
    public partial class ChooseProblemSheet
    {
        public Range orig_range;
        public string testcase_id;
        private void ChooseProblemSheet_Startup(object sender, System.EventArgs e)
        {
        }

        private void ChooseProblemSheet_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO 设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.buttonProblemSubmit.Click += new System.EventHandler(this.buttonProblemSubmit_Click);
            this.checkBoxChoose.CheckedChanged += new System.EventHandler(this.checkBoxChoose_CheckedChanged);
            this.SelectionChange += new Microsoft.Office.Interop.Excel.DocEvents_SelectionChangeEventHandler(this.ChooseProblemSheet_SelectionChange);
            this.Deactivate += new Microsoft.Office.Interop.Excel.DocEvents_DeactivateEventHandler(this.ChooseProblemSheet_Deactivate);
            this.Startup += new System.EventHandler(this.ChooseProblemSheet_Startup);
            this.Shutdown += new System.EventHandler(this.ChooseProblemSheet_Shutdown);

        }

        #endregion

        private void buttonProblemSubmit_Click(object sender, EventArgs e)
        {
            Visible = XlSheetVisibility.xlSheetHidden;
            Globals.TestRecordSheet.Visible = XlSheetVisibility.xlSheetVisible;
            Globals.TestRecordSheet.Activate();
        }

        private void ChooseProblemSheet_SelectionChange(Range Target)
        {
            int max_row = UsedRange.Rows.Count;
            if (Target.Row > max_row || Target.Row == 1) {
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
                else {
                    checkBoxChoose.Checked = false;
                }
            }
            else {
                checkBoxChoose.Visible = false;
            }
            buttonProblemSubmit.Top = Cells[Target.Row, 4].Top;
            buttonProblemSubmit.Left = Cells[Target.Row, 4].Left;
            buttonProblemSubmit.Visible = true;
        }

        private void checkBoxChoose_CheckedChanged(object sender, EventArgs e)
        {
            int problem_id_column = RangeUtils.get_column_by_title(Cells, UsedRange.Columns.Count,"问题标识");            
            Range range = checkBoxChoose.Tag as Range;
            String problem_id = Cells[range.Row, problem_id_column].Text;
            if (checkBoxChoose.Checked)
            {
                range.Value = "是";
                Globals.TestRecordAndProblemSheet.set_testcase_and_problem_cell(testcase_id, problem_id, "是");
            }
            else {
                range.Value = "";
                Globals.TestRecordAndProblemSheet.set_testcase_and_problem_cell(testcase_id, problem_id, "");
            }

            orig_range.Value = Globals.TestRecordAndProblemSheet.get_problems_by_testcase_id(testcase_id);
        }

        private void ChooseProblemSheet_Deactivate()
        {
            checkBoxChoose.Visible = false;
            buttonProblemSubmit.Visible = false;
            Visible = XlSheetVisibility.xlSheetHidden;
            Globals.TestRecordSheet.Visible = XlSheetVisibility.xlSheetVisible;
            Globals.TestRecordSheet.Activate();
        }
    }
}
