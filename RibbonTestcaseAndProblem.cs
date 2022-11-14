using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.IO;

namespace TestCaseAndProblem
{
    public partial class RibbonTestcaseAndProblem
    {
        private ActionsPaneControlTestCase actionsPaneControlTestCase; 
        private void RibbonTestcaseAndProblem_Load(object sender, RibbonUIEventArgs e)
        {
            actionsPaneControlTestCase = Globals.ThisWorkbook.ActionsPaneControlTestCase;
            Globals.ThisWorkbook.ActionsPane.Controls.Add(actionsPaneControlTestCase);
            checkBoxToggleActionsPane.Checked = true;
        }

        private void buttonGenTestCaseWord_Click(object sender, RibbonControlEventArgs e)
        {
            backgroundWorkerTestRecord.RunWorkerAsync("产生测试记录");
        }

        private void gen_testRecord()
        {
            String dir_name = Globals.EditItemsSheet.get_file_save_dir();
            String project_name = editBoxProjectName.Text;
            String testcase_template_file_name = Globals.EditItemsSheet.get_testcase_template_file_name();

            buttonGenTestCaseWord.Enabled = false;
            OfficeWordUtils.createWordTestRecordByDocmentTemplate(testcase_template_file_name, dir_name + project_name + "测试记录.docx", Globals.TestRecordSheet.Cells, backgroundWorkerTestRecord);
            OfficeWordUtils.openWordFile(dir_name + project_name + "测试记录.docx");
            buttonGenTestCaseWord.Enabled = true;
        }

        private void buttonGenProblemWord_Click(object sender, RibbonControlEventArgs e)
        {
            backgroundWorkerTestRecord.RunWorkerAsync("产生问题报告");
        }

        private void genProblemWord()
        {
            String dir_name = Globals.EditItemsSheet.get_file_save_dir();
            String project_name = editBoxProjectName.Text;
            String problem_template_file_name = Globals.EditItemsSheet.get_problem_template_file_name();
            
            buttonGenProblemWord.Enabled = false;
            OfficeWordUtils.createWordProblemByDocmentTemplate(problem_template_file_name, dir_name + project_name + "问题报告单.docx", Globals.ProblemSheet.Cells, backgroundWorkerTestRecord);
            OfficeWordUtils.openWordFile(dir_name + project_name + "问题报告单.docx");
            buttonGenProblemWord.Enabled = true;
        }

        private void checkBoxSetting_Click(object sender, RibbonControlEventArgs e)
        {
            if (checkBoxSetting.Checked)
            {
                Globals.EditItemsSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                Globals.EditItemsSheet.Activate();
            }
            else
            {
                Globals.EditItemsSheet.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            }
        }

        private void checkBoxToggleActionsPane_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisWorkbook.Application.DisplayDocumentActionTaskPane = checkBoxToggleActionsPane.Checked;
        }

        private void backgroundWorkerTestRecord_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            try
            {
                String argument = e.Argument as String;
                if (String.Equals(argument, "产生测试记录"))
                {
                    gen_testRecord();
                }
                else if (String.Equals(argument, "产生问题报告"))
                {
                    genProblemWord();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("生成错误", ex.ToString());
                throw ex;
            }
        }

        private void backgroundWorkerTestRecord_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                Type errorType = e.Error.GetType();
                switch (errorType.Name)
                {
                    case "ArgumentNullException":
                    case "MyException":
                        //do something.
                        break;
                    default:
                        //do something.
                        break;
                }
                MessageBox.Show("生成错误", e.Error.Message);
                buttonGenTestCaseWord.Enabled = true;
                buttonGenTestCaseWord.Enabled = true;
                notifyIconReport.ShowBalloonTip(1000, "生成进度", "生成错误", ToolTipIcon.Info);
            }
            else {
                buttonGenTestCaseWord.Enabled = true;
                buttonGenTestCaseWord.Enabled = true;
                notifyIconReport.ShowBalloonTip(1000, "生成进度", "完成", ToolTipIcon.Info);
                editBoxLog.Text = "完成";
                //actionsPaneControlTestCase.show_progress(100, "完成");
            }
        }

        private void backgroundWorkerTestRecord_ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e)
        {
            //notifyIconReport.ShowBalloonTip(1000, "生成进度", "完成" +e.ProgressPercentage, ToolTipIcon.Info);
            //actionsPaneControlTestCase.show_progress(e.ProgressPercentage, e.UserState.ToString());
            editBoxLog.Text = e.UserState.ToString();
        }
    }
}
