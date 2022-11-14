using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Tools.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace TestCaseAndProblem
{
    public partial class EditProblemSheet
    {
        public int row_in_problems = 1;
        private void EditProblemSheet_Startup(object sender, System.EventArgs e)
        {
        }

        private void EditProblemSheet_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO 设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.buttonSaveProblems.Click += new System.EventHandler(this.buttonSaveProblems_Click);
            this.buttonInsertPicture.Click += new System.EventHandler(this.buttonInsertPicture_Click);
            this.buttonInsertPicture.MouseHover += new System.EventHandler(this.buttonInsertPicture_MouseHover);
            this.SelectionChange += new Microsoft.Office.Interop.Excel.DocEvents_SelectionChangeEventHandler(this.EditProblemSheet_SelectionChange);
            this.ActivateEvent += new Microsoft.Office.Interop.Excel.DocEvents_ActivateEventHandler(this.EditProblemSheet_ActivateEvent);
            this.Deactivate += new Microsoft.Office.Interop.Excel.DocEvents_DeactivateEventHandler(this.EditProblemSheet_Deactivate);
            this.Startup += new System.EventHandler(this.EditProblemSheet_Startup);
            this.Shutdown += new System.EventHandler(this.EditProblemSheet_Shutdown);

        }

        #endregion

        private void EditProblemSheet_SelectionChange(Excel.Range Target)
        {
            if (Target.Row > 1)
            {
                int iLast = UsedRange.Columns.Count;
                if (iLast > 1)
                {
                    buttonSaveProblems.Top = Cells[Target.Row, iLast + 1].top;
                    buttonSaveProblems.Left = Cells[Target.Row, iLast + 1].Left;
                    buttonSaveProblems.Visible = true;
                }
            }
            else
            {
                buttonSaveProblems.Visible = false;
            }

            String title_str = Cells[1, Target.Column].Text;
            if (Target.Row > 1 && !String.IsNullOrEmpty(title_str) && title_str.Contains("图片"))//问题图片
            {
                if (String.IsNullOrEmpty(Target.Text))
                {
                    buttonInsertPicture.Top = Target.Top;
                    buttonInsertPicture.Left = Target.Left;
                    buttonInsertPicture.Tag = Target;
                    buttonInsertPicture.Visible = true;
                    pictureBoxPreview.Visible = false;
                }
                else
                {
                    pictureBoxPreview.Top = Target.Top;
                    pictureBoxPreview.Left = Target.Left;
                    try
                    {
                        Image img = PictureUtils.getImage(Target.Text);
                        pictureBoxPreview.Height = img.Height * 72 / 96;
                        pictureBoxPreview.Width = img.Width * 72 / 96;
                        pictureBoxPreview.Image = img;
                        pictureBoxPreview.Visible = true;
                        buttonInsertPicture.Visible = false;
                    }
                    catch (Exception)
                    {
                        pictureBoxPreview.Image = null;
                        pictureBoxPreview.Visible = true;
                        buttonInsertPicture.Visible = false;
                    }
                }
            }
            else
            {
                buttonInsertPicture.Visible = false;
                pictureBoxPreview.Visible = false;
            }
        }

        public void save_result()
        {
            int problemSheet_last_column = Globals.ProblemSheet.UsedRange.Columns.Count;
            int editProblemSheet_last_column = UsedRange.Columns.Count;
            for (int i = 1; i <= problemSheet_last_column; i++)
            {
                for (int j = 1; j <= editProblemSheet_last_column; j++)
                {
                    String et = Cells[1, j].Text;
                    String ct = Globals.ProblemSheet.Cells[1, i].Text;
                    if (!String.IsNullOrEmpty(et) && !String.IsNullOrEmpty(ct) && String.Equals(et, ct))
                    {
                        String result = "";
                        int column_max_row = UsedRange.Rows.Count;
                        for (int k = 2; k <= column_max_row; k++)
                        {
                            result = result + (k - 1) + "." + Cells[k, j].Text + "\r\n";
                        }
                        Globals.ProblemSheet.Cells[row_in_problems, i].Value = result;
                    }
                }
            }
        }
        private void buttonSaveProblems_Click(object sender, EventArgs e)
        {
            save_result();
            Visible = Excel.XlSheetVisibility.xlSheetHidden;
            Globals.ProblemSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
            Globals.ProblemSheet.Activate();
        }

        private void buttonInsertPicture_Click(object sender, EventArgs e)
        {
            String guid_str = PictureUtils.insert_picture();
            Excel.Range range = buttonInsertPicture.Tag as Excel.Range;
            range.Value = guid_str;
        }

        private void buttonInsertPicture_MouseHover(object sender, EventArgs e)
        {
            var toolTip1 = new ToolTip();

            toolTip1.AutoPopDelay = 10000;
            toolTip1.InitialDelay = 500;
            toolTip1.ReshowDelay = 500;

            toolTip1.ShowAlways = true;

            toolTip1.SetToolTip(buttonInsertPicture, @"截图软件截取内容后，点击我插入图片");
        }

        private void EditTestCaseSheet_ActivateEvent()
        {
            UsedRange.Cells.Interior.Color = System.Drawing.Color.Yellow.ToArgb();
            UsedRange.Borders.LineStyle = 1;
            UsedRange.Interior.Color = System.Drawing.Color.White.ToArgb();
            UsedRange.Font.Color = System.Drawing.Color.White.ToArgb();
        }

        private void EditProblemSheet_ActivateEvent()
        {
            UsedRange.Cells.Interior.Color = System.Drawing.Color.Yellow.ToArgb();
            UsedRange.Borders.LineStyle = 1;
            UsedRange.Interior.Color = System.Drawing.Color.White.ToArgb();
            UsedRange.Font.Color = System.Drawing.Color.White.ToArgb();
        }

        private void EditProblemSheet_Deactivate()
        {
            buttonInsertPicture.Visible = false;
            buttonSaveProblems.Visible = false;
            Visible = Excel.XlSheetVisibility.xlSheetHidden;
            Globals.ProblemSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
            Globals.ProblemSheet.Activate();
        }
    }
}
