using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace TestCaseAndProblem
{
    public partial class TestRecordSheet : InterfaceUpdateText
    {
        private string relate_problem_title;
        private string testcase_pass_title;
        private int testcase_name_column;
        private int testcase_id_column;
        private void hilight_detailEditItems() {
            Range rangeDetail = null;
            Range rangeOther = null;
            String detail_edit_columns = Globals.EditItemsSheet.get_detail_edit_item_testcase();
            foreach (Range range in UsedRange.Columns)
            {
                String ct = Cells[1, range.Column].Text;
                if (detail_edit_columns.Contains(ct))
                {

                    if (rangeDetail == null)
                    {
                        rangeDetail = range;
                    }
                    else
                    {
                        rangeDetail = Application.Union(rangeDetail, range);
                    }
                }
                else
                {
                    if (rangeOther == null)
                    {
                        rangeOther = range;
                    }
                    else
                    {
                        rangeOther = Application.Union(rangeOther, range);
                    }
                }
            }
            if (rangeDetail != null)
            {
                rangeDetail.Cells.Interior.Color = System.Drawing.Color.RosyBrown.ToArgb();
                rangeDetail.Borders.LineStyle = 1;
                rangeDetail.Interior.ColorIndex = 39;
                rangeDetail.Font.Color = System.Drawing.Color.Red.ToArgb();
            }

            if (rangeOther != null)
            {
                rangeOther.Cells.Interior.Color = System.Drawing.Color.Yellow.ToArgb();
                rangeOther.Borders.LineStyle = 1;
                rangeOther.Interior.Color = System.Drawing.Color.White.ToArgb();
                rangeOther.Font.Color = System.Drawing.Color.White.ToArgb();
            }

            char A = (char)('A' + UsedRange.Columns.Count - 1);
            Range rangeFirstRow = get_Range("A1:" + A + "1");
            rangeFirstRow.Cells.Interior.Color = System.Drawing.Color.Red.ToArgb();
            rangeFirstRow.Borders.LineStyle = 1;
            rangeFirstRow.Interior.Color = System.Drawing.Color.LightBlue.ToArgb();
            rangeFirstRow.Font.Color = System.Drawing.Color.Red.ToArgb();
        }

        private void hilight_detailEditItems_row(int row)
        {
            Range rangeDetail = null;
            Range rangeOther = null;
            String detail_edit_columns = Globals.EditItemsSheet.get_detail_edit_item_testcase();
            bool has_text = false;

            int max_column = UsedRange.Columns.Count;
            for (int col = 1; col <= max_column; col++)
            {
                String ct = Cells[1, col].Text;
                if (detail_edit_columns.Contains(ct))
                {

                    if (rangeDetail == null)
                    {
                        rangeDetail = Cells[row, col];
                    }
                    else
                    {
                        rangeDetail = Application.Union(rangeDetail, Cells[row, col]);
                    }
                }
                else
                {
                    if (rangeOther == null)
                    {
                        rangeOther = Cells[row, col];
                    }
                    else
                    {
                        rangeOther = Application.Union(rangeOther, Cells[row, col]);
                    }
                }
                if (!String.IsNullOrWhiteSpace(Cells[row, col].Text))
                {
                    has_text = true;
                }
            }
            if (rangeDetail != null)
            {
                if (has_text)
                {
                    rangeDetail.Cells.Interior.Color = System.Drawing.Color.RosyBrown.ToArgb();
                    rangeDetail.Borders.LineStyle = 1;
                    rangeDetail.Interior.ColorIndex = 39;
                    rangeDetail.Font.Color = System.Drawing.Color.Red.ToArgb();
                }
                else
                {
                    rangeDetail.Cells.Interior.Color = System.Drawing.Color.White.ToArgb();
                    rangeDetail.Borders.LineStyle = 0;
                    rangeDetail.Interior.ColorIndex = 0;
                    rangeDetail.Font.Color = System.Drawing.Color.Black.ToArgb();
                }
            }

            if (rangeOther != null)
            {
                if (has_text)
                {
                    rangeOther.Cells.Interior.Color = System.Drawing.Color.Yellow.ToArgb();
                    rangeOther.Borders.LineStyle = 1;
                    rangeOther.Interior.Color = System.Drawing.Color.White.ToArgb();
                    rangeOther.Font.Color = System.Drawing.Color.White.ToArgb();
                }
                else
                {
                    rangeOther.Cells.Interior.Color = System.Drawing.Color.White.ToArgb();
                    rangeOther.Borders.LineStyle = 0;
                    rangeOther.Interior.ColorIndex = 0;
                    rangeOther.Font.Color = System.Drawing.Color.Black.ToArgb();
                }
            }
        }

        private void TestCases_Startup(object sender, System.EventArgs e)
        {
            hilight_detailEditItems();
            relate_problem_title = Globals.EditItemsSheet.get_testcase_relate_problem_title();
            testcase_pass_title = Globals.EditItemsSheet.get_testcase_pass_title();
            testcase_id_column = RangeUtils.get_column_by_title(Cells, UsedRange.Columns.Count, "用例标识");
            testcase_name_column = RangeUtils.get_column_by_title(Cells, UsedRange.Columns.Count, "用例名称");
            set_comboBoxTestor();
        }

        private void set_comboBoxTestor()
        {
            comboBoxTestor.Items.Clear();
            foreach (String v in Globals.EditItemsSheet.get_testor())
            {
                comboBoxTestor.Items.Add(v);
            }
        }

        private void TestCases_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO 设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.buttonEditTestCase.Click += new System.EventHandler(this.buttonEditTestCase_Click);
            this.buttonChooseProblem.Click += new System.EventHandler(this.buttonChooseProblem_Click);
            this.comboBoxPassFail.SelectedIndexChanged += new System.EventHandler(this.comboBoxPassFail_SelectedIndexChanged);
            this.comboBoxTestor.SelectedIndexChanged += new System.EventHandler(this.comboBoxTestor_SelectedIndexChanged);
            this.SelectionChange += new Microsoft.Office.Interop.Excel.DocEvents_SelectionChangeEventHandler(this.TestCasesSheet_SelectionChange);
            this.ActivateEvent += new Microsoft.Office.Interop.Excel.DocEvents_ActivateEventHandler(this.TestCasesSheet_ActivateEvent);
            this.Deactivate += new Microsoft.Office.Interop.Excel.DocEvents_DeactivateEventHandler(this.TestCasesSheet_Deactivate);
            this.Startup += new System.EventHandler(this.TestCases_Startup);
            this.Shutdown += new System.EventHandler(this.TestCases_Shutdown);
            this.Change += new Microsoft.Office.Interop.Excel.DocEvents_ChangeEventHandler(this.TestRecordSheet_Change);

        }
        #endregion

        private void get_all_step(string title, string string_steps, int dest_column)
        {
            int dest_row = 1;
            Globals.EditTestRecordSheet.Cells[dest_row++, dest_column] = title;
            
            foreach (Match m in Regex.Matches(string_steps, @"^(.+)$", RegexOptions.Multiline))
            {
                String str = Regex.Replace(m.Value, @"^\d+\.\s*", "");
                Globals.EditTestRecordSheet.Cells[dest_row++, dest_column] = Regex.Replace(str, @"[\r\n]*", ""); 
            }
        }

        private void TestCasesSheet_SelectionChange(Excel.Range Target)
        {
            String detail_edit_columns = Globals.EditItemsSheet.get_detail_edit_item_testcase();
            if (Target.Row > 1 &&detail_edit_columns.Contains(Cells[1, Target.Column].Text))
            {
                buttonEditTestCase.Top = Cells[Target.Row, Target.Column + 1].top;
                buttonEditTestCase.Left = Cells[Target.Row, Target.Column + 1].Left;
                buttonEditTestCase.Tag = Target;
                buttonEditTestCase.Visible = true;
            }
            else {
                buttonEditTestCase.Visible = false;
            }        

            if (Target.Row > 1 && String.Equals(Cells[1, Target.Column].Text, relate_problem_title))
            {
                buttonChooseProblem.Top = Target.Top;
                buttonChooseProblem.Left = Target.Left;
                buttonChooseProblem.Tag = Target;
                buttonChooseProblem.Visible = true;
            }
            else
            {
                buttonChooseProblem.Visible = false;
            }

            if (Target.Row > 1 && String.Equals(Cells[1, Target.Column].Text, testcase_pass_title))
            {
                comboBoxPassFail.Top = Target.Top;
                comboBoxPassFail.Left = Target.Left;
                comboBoxPassFail.Tag = Target;
                comboBoxPassFail.Visible = true;
            }
            else
            {
                comboBoxPassFail.Visible = false;
            }

            if (Target.Row > 1 && String.Equals(Cells[1, Target.Column].Text, "测试人员"))
            {
                comboBoxTestor.Top = Target.Top;
                comboBoxTestor.Left = Target.Left;
                comboBoxTestor.Tag = Target;
                comboBoxTestor.Visible = true;
            }
            else
            {
                comboBoxTestor.Visible = false;
            }
            Globals.ThisWorkbook.ActionsPaneControlTestCase.set_range(Target);
        }

        private void buttonEditTestCase_Click(object sender, EventArgs e)
        {
            int dest_column = 1;
            Excel.Range range_button = buttonEditTestCase.Tag as Excel.Range;
            int src_row = range_button.Row;
            String detail_edit_columns = Globals.EditItemsSheet.get_detail_edit_item_testcase();
            Globals.EditTestRecordSheet.UsedRange.Clear();
            foreach (Range range in UsedRange.Columns)
            {
                String ct = Cells[1, range.Column].Text;
                if (detail_edit_columns.Contains(ct))
                {
                    get_all_step(ct, Cells[src_row, range.Column].Text, dest_column++);
                }
            }
            Globals.EditTestRecordSheet.row_in_testcases = src_row;
            Globals.EditTestRecordSheet.Visible = XlSheetVisibility.xlSheetVisible;
            Globals.EditTestRecordSheet.Activate();
        }

        private void buttonChooseProblem_Click(object sender, EventArgs e)
        {
            int max_source_row = Globals.ProblemSheet.UsedRange.Rows.Count;
            int max_source_column = Globals.ProblemSheet.UsedRange.Columns.Count;
            int testcase_id_column = RangeUtils.get_column_by_title(Cells, UsedRange.Columns.Count, "用例标识");
            if (testcase_id_column == 0)
            {
                return;
            }
            Globals.ChooseProblemSheet.UsedRange.Clear();
            Globals.ChooseProblemSheet.Cells[1, 3].Value = "选择";
            int dest_column = 1;
            String problem_string = (buttonChooseProblem.Tag as Range).Text;
            for (int column = 1; column <= max_source_column; column++) {
                String pt = Globals.ProblemSheet.Cells[1, column].Text;
                String et1 = "问题名称";
                String et2 = "问题标识";

                if (String.Equals(pt, et1) ||String.Equals(pt, et2))
                {                   
                    for (int row = 1; row <= max_source_row; row++) {
                        Globals.ChooseProblemSheet.Cells[row, dest_column].Value = Globals.ProblemSheet.Cells[row, column].Value;
                        if (problem_string.Contains(Globals.ProblemSheet.Cells[row, column].Text)) {
                            Globals.ChooseProblemSheet.Cells[row, 3].Value = "是";
                        }
                    }
                    dest_column++;
                }                
            }
            Globals.ChooseProblemSheet.orig_range = buttonChooseProblem.Tag as Range;
            
            String testcase_id = Globals.TestRecordSheet.Cells[Globals.ChooseProblemSheet.orig_range.Row, testcase_id_column].Text;
            Globals.ChooseProblemSheet.testcase_id = testcase_id;
            Globals.ChooseProblemSheet.Visible = XlSheetVisibility.xlSheetVisible;
            Globals.ChooseProblemSheet.Activate();            
        }

        private void TestCasesSheet_ActivateEvent()
        {
            set_comboBoxTestor();
            Globals.ThisWorkbook.ActionsPaneControlTestCase.InterfaceUpdateText = this;
        }

        public String get_testcase_name_by_id(String testcase_id) {
            int row = RangeUtils.get_row_by_title(Cells, testcase_id, testcase_id_column);
            if (row == 0) {
                return "";
            }
            return Cells[row, testcase_name_column].Text;
        }

        private void TestCasesSheet_Deactivate()
        {
        }

        public void selelect_lost_focus()
        {
            
        }

        public void selelect_focus(string text)
        {
        }

        private void comboBoxPassFail_SelectedIndexChanged(object sender, EventArgs e)
        {
            Range range = comboBoxPassFail.Tag as Range;
            if(range != null)
            {
                range.Value = comboBoxPassFail.Text;
            }            
        }

        private void comboBoxTestor_SelectedIndexChanged(object sender, EventArgs e)
        {
            Range range = comboBoxTestor.Tag as Range;
            if (range != null)
            {
                range.Value = comboBoxTestor.Text;
            }
        }

        private void TestRecordSheet_Change(Range Target)
        {
            hilight_detailEditItems_row(Target.Row);
        }
    }
}
