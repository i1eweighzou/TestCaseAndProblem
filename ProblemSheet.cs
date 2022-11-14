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
    public partial class ProblemSheet:InterfaceUpdateText
    {
        private string relate_testcase_title;
        private string problem_type_title;
        private string problem_level_title;
        
        private void hilight_detailEditItems()
        {
            Range rangeDetail = null;
            Range rangeOther = null;
            String detail_edit_columns = Globals.EditItemsSheet.get_detail_edit_item_problem();
            foreach (Range range in UsedRange.Columns)
            {
                String ct = Cells[1, range.Column].Text;
                if (detail_edit_columns.Contains(ct))
                {
                    
                    if (rangeDetail == null)
                    {
                        rangeDetail = range;
                    }
                    else {
                        rangeDetail = Application.Union(rangeDetail, range);
                    }                    
                }
                else {                    
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

            char A = (char)('A' + UsedRange.Columns.Count - 1) ;
            Range rangeFirstRow = get_Range("A1:"+A+"1");          
            rangeFirstRow.Cells.Interior.Color = System.Drawing.Color.Red.ToArgb();
            rangeFirstRow.Borders.LineStyle = 1;
            rangeFirstRow.Interior.Color = System.Drawing.Color.LightBlue.ToArgb();
            rangeFirstRow.Font.Color = System.Drawing.Color.Red.ToArgb();
        }

        private void hilight_detailEditItems_row(int row)
        {
            Range rangeDetail = null;
            Range rangeOther = null;
            String detail_edit_columns = Globals.EditItemsSheet.get_detail_edit_item_problem();
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
                if (!String.IsNullOrWhiteSpace(Cells[row, col].Text)) {
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
                else {
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

        private void ProblemSheet_Startup(object sender, System.EventArgs e)
        {
            hilight_detailEditItems();
            relate_testcase_title = Globals.EditItemsSheet.get_problem_relate_testcase_title();
            problem_type_title = Globals.EditItemsSheet.get_problem_type_title();
            problem_level_title = Globals.EditItemsSheet.get_problem_level_title();
            set_comboBoxReportor();
        }

        private void set_comboBoxReportor()
        {
            comboBoxReportor.Items.Clear();
            foreach (String v in Globals.EditItemsSheet.get_reportor())
            {
                comboBoxReportor.Items.Add(v);
            }
        }

        private void ProblemSheet_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO 设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.buttonEditProblem.Click += new System.EventHandler(this.buttonEditProblem_Click);
            this.comboBoxProblemType.SelectedIndexChanged += new System.EventHandler(this.comboBoxProblemType_SelectedIndexChanged);
            this.comboBoxProblemLevel.SelectedIndexChanged += new System.EventHandler(this.comboBoxProblemLevel_SelectedIndexChanged);
            this.comboBoxReportor.SelectedIndexChanged += new System.EventHandler(this.comboBoxReportor_SelectedIndexChanged);
            this.SelectionChange += new Microsoft.Office.Interop.Excel.DocEvents_SelectionChangeEventHandler(this.ProblemSheet_SelectionChange);
            this.ActivateEvent += new Microsoft.Office.Interop.Excel.DocEvents_ActivateEventHandler(this.ProblemSheet_ActivateEvent);
            this.Deactivate += new Microsoft.Office.Interop.Excel.DocEvents_DeactivateEventHandler(this.ProblemSheet_Deactivate);
            this.Startup += new System.EventHandler(this.ProblemSheet_Startup);
            this.Shutdown += new System.EventHandler(this.ProblemSheet_Shutdown);
            this.Change += new Microsoft.Office.Interop.Excel.DocEvents_ChangeEventHandler(this.ProblemSheet_Change);

        }

        #endregion

        private void ProblemSheet_SelectionChange(Excel.Range Target)
        {
            String detail_edit_columns = Globals.EditItemsSheet.get_detail_edit_item_problem();
            if (Target.Row > 1 && detail_edit_columns.Contains(Cells[1, Target.Column].Text))
            {
                buttonEditProblem.Top = Cells[Target.Row, Target.Column + 1].top;
                buttonEditProblem.Left = Cells[Target.Row, Target.Column + 1].Left;
                buttonEditProblem.Tag = Target;
                buttonEditProblem.Visible = true;
            }
            else
            {
                buttonEditProblem.Visible = false;
            }

            if (Target.Row > 1 && String.Equals(Cells[1, Target.Column].Text, problem_type_title))
            {
                comboBoxProblemType.Top = Target.Top;
                comboBoxProblemType.Left = Target.Left;
                comboBoxProblemType.Tag = Target;
                comboBoxProblemType.Visible = true;
            }
            else
            {
                comboBoxProblemType.Visible = false;
            }

            if (Target.Row > 1 && String.Equals(Cells[1, Target.Column].Text, problem_level_title))
            {
                comboBoxProblemLevel.Top = Target.Top;
                comboBoxProblemLevel.Left = Target.Left;
                comboBoxProblemLevel.Tag = Target;
                comboBoxProblemLevel.Visible = true;
            }
            else
            {
                comboBoxProblemLevel.Visible = false;
            }

            if (Target.Row > 1 && String.Equals(Cells[1, Target.Column].Text, "报告人"))
            {
                comboBoxReportor.Top = Target.Top;
                comboBoxReportor.Left = Target.Left;
                comboBoxReportor.Tag = Target;
                comboBoxReportor.Visible = true;
            }
            else
            {
                comboBoxReportor.Visible = false;
            }

            Globals.ThisWorkbook.ActionsPaneControlTestCase.set_range(Target);
        }

        private void get_all_step(string title, string string_steps, int dest_column)
        {
            int dest_row = 1;
            Globals.EditProblemSheet.Cells[dest_row++, dest_column] = title;

            foreach (Match m in Regex.Matches(string_steps, @"^(.+)$", RegexOptions.Multiline))
            {
                String str = Regex.Replace(m.Value, @"^\d+\.\s*", "");
                Globals.EditProblemSheet.Cells[dest_row++, dest_column] = Regex.Replace(str, @"[\r\n]*", "");
            }
        }

        private void buttonEditProblem_Click(object sender, EventArgs e)
        {
            int dest_column = 1;
            Excel.Range range_button = buttonEditProblem.Tag as Excel.Range;
            int src_row = range_button.Row;
            String detail_edit_columns = Globals.EditItemsSheet.get_detail_edit_item_problem();
            Globals.EditProblemSheet.UsedRange.Clear();
            foreach (Range range in UsedRange.Columns)
            {
                String ct = Cells[1, range.Column].Text;
                if (detail_edit_columns.Contains(ct))
                {
                    get_all_step(ct, Cells[src_row, range.Column].Text, dest_column++);
                }
            }
            Globals.EditProblemSheet.row_in_problems = src_row;
            Globals.EditProblemSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
            Globals.EditProblemSheet.Activate();
        }       

        private void ProblemSheet_ActivateEvent()
        {
            set_comboBoxReportor();
            Globals.ThisWorkbook.ActionsPaneControlTestCase.InterfaceUpdateText = this;
        }

        private void ProblemSheet_Deactivate()
        {
        }

        public void selelect_lost_focus()
        {
        }

        public void selelect_focus(string text)
        {
        }

        private void comboBoxProblemType_SelectedIndexChanged(object sender, EventArgs e)
        {
            Range range = comboBoxProblemType.Tag as Range;
            if (range != null)
            {
                range.Value = comboBoxProblemType.Text;
            }
        }

        private void comboBoxProblemLevel_SelectedIndexChanged(object sender, EventArgs e)
        {
            Range range = comboBoxProblemLevel.Tag as Range;
            if (range != null)
            {
                range.Value = comboBoxProblemLevel.Text;
            }
        }

        private void comboBoxReportor_SelectedIndexChanged(object sender, EventArgs e)
        {
            Range range = comboBoxReportor.Tag as Range;
            if (range != null)
            {
                range.Value = comboBoxReportor.Text;
            }
        }

        private void ProblemSheet_Change(Range Target)
        {
            hilight_detailEditItems_row(Target.Row);
        }
    }
}
