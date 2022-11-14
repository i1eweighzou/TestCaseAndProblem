using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
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
    public partial class EditItemsSheet
    {
        private void hilight_and_lock_first_column()
        {            
            int max_row = UsedRange.Rows.Count;
            Range rangeFirstColumn = Range["A1:A" + max_row];
            rangeFirstColumn.Cells.Interior.Color = System.Drawing.Color.Red.ToArgb();
            rangeFirstColumn.Borders.LineStyle = 1;
            rangeFirstColumn.Interior.Color = System.Drawing.Color.LightBlue.ToArgb();
            rangeFirstColumn.Font.Color = System.Drawing.Color.Red.ToArgb();
            rangeFirstColumn.Locked = true;
        }
    
        public String get_detail_edit_item_testcase()
        {
            int row = RangeUtils.get_row_by_title(UsedRange, "测试记录详细编辑项");
            String ret = "";
            for (int column = 2; column <= UsedRange.Columns.Count; column++) {
                if(String.IsNullOrEmpty(UsedRange[row, column].Text)) {
                    break;
                }
                if (String.IsNullOrEmpty(ret))
                {
                    ret = UsedRange[row, column].Text;
                }
                else { 
                    ret += "\n"+ UsedRange[row, column].Text;
                }
            }
            return ret;
        }

        public int get_detail_edit_item_testcase_count()
        {
            int row = RangeUtils.get_row_by_title(UsedRange, "测试记录详细编辑项");
            int count=0;
            for (int column = 2; column <= UsedRange.Columns.Count; column++)
            {
                if (String.IsNullOrEmpty(UsedRange[row, column].Text))
                {
                    break;
                }
                count++;
            }
            return count;
        }

        public String get_detail_edit_item_problem()
        {
            int row = RangeUtils.get_row_by_title(UsedRange, "问题报告详细编辑项");
            String ret = "";
            for (int column = 2; column <= UsedRange.Columns.Count; column++)
            {
                if (String.IsNullOrEmpty(UsedRange[row, column].Text)) {
                    break;
                }
                if (String.IsNullOrEmpty(ret))
                {
                    ret = UsedRange[row, column].Text;
                }
                else
                {
                    ret += "\n" + UsedRange[row, column].Text;
                }
            }
            return ret;
        }

        public int get_detail_edit_item_problem_count()
        {
            int row = RangeUtils.get_row_by_title(UsedRange, "问题报告详细编辑项");
            int count = 0;
            for (int column = 2; column <= UsedRange.Columns.Count; column++)
            {
                if (String.IsNullOrEmpty(UsedRange[row, column].Text))
                {
                    break;
                }
                count++;
            }
            return count;
        }

        public String get_file_save_dir() {
            int row = RangeUtils.get_row_by_title(UsedRange, "文件保存路径");
            String content = Cells[row, 2].Text;
            content = content.Trim();
            string dir = content + @"\";
            if (!Directory.Exists(dir))
            {
                try {
                    Directory.CreateDirectory(dir);
                }
                catch {
                    dir = @"C:\TestcaseAndRecord";
                    Cells[row, 2].Value = dir;
                    dir = dir + @"\";
                    Directory.CreateDirectory(dir);
                }                
            }
            return dir;
        }

        public String get_testcase_template_file_name()
        {
            int row = RangeUtils.get_row_by_title(UsedRange, "测试记录模板");
            String content = Cells[row, 2].Text;
            content = content.Trim();
            String dir = get_file_save_dir() + @"template\";
            Directory.CreateDirectory(dir);
            String testcase_template_file_name_full = dir  + content;
            try
            {
                if (!File.Exists(testcase_template_file_name_full))
                {
                    FileStream fileStream = new FileStream(testcase_template_file_name_full, FileMode.OpenOrCreate);
                    fileStream.Write(Properties.Resources.test_record, 0, Properties.Resources.test_record.Length);
                    fileStream.Close();
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.ToString(), "创建失败");
            }
            return testcase_template_file_name_full;
        }

        public String get_problem_template_file_name()
        {
            int row = RangeUtils.get_row_by_title(UsedRange, "问题报告模板");
            String content = Cells[row, 2].Text;
            content = content.Trim();
            String dir = get_file_save_dir() + @"template\";
            Directory.CreateDirectory(dir);
            String problem_template_file_name_full = dir + content;
            try
            {
                if (!File.Exists(problem_template_file_name_full))
                {
                    FileStream fileStream = new FileStream(problem_template_file_name_full, FileMode.OpenOrCreate);
                    fileStream.Write(Properties.Resources.problem, 0, Properties.Resources.problem.Length);
                    fileStream.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "创建失败");
            }
            return problem_template_file_name_full;
        }

        public String get_detail_edit_item_testcase_template()
        {
            int row = RangeUtils.get_row_by_title(UsedRange, "测试记录模板详细编辑项");
            String ret = "";
            for (int column = 2; column <= UsedRange.Columns.Count; column++)
            {
                if (String.IsNullOrEmpty(UsedRange[row, column].Text))
                {
                    break;
                }
                if (String.IsNullOrEmpty(ret))
                {
                    ret = UsedRange[row, column].Text;
                }
                else
                {
                    ret += "\n" + UsedRange[row, column].Text;
                }
            }
            return ret;
        }

        public String get_detail_edit_item_problem_template()
        {
            int row = RangeUtils.get_row_by_title(UsedRange, "问题报告模板详细编辑项");
            String ret = "";
            for (int column = 2; column <= UsedRange.Columns.Count; column++)
            {
                if (String.IsNullOrEmpty(UsedRange[row, column].Text))
                {
                    break;
                }
                if (String.IsNullOrEmpty(ret))
                {
                    ret = UsedRange[row, column].Text;
                }
                else
                {
                    ret += "\n" + UsedRange[row, column].Text;
                }
            }
            return ret;
        }

        public String get_outline_level_title(int level) {
            int row = 0;
            switch (level) {
                case 1:
                    row = RangeUtils.get_row_by_title(UsedRange, "一级目录");
                    break;
                case 2:
                    row = RangeUtils.get_row_by_title(UsedRange, "二级目录");
                    break;
                case 3:
                    row = RangeUtils.get_row_by_title(UsedRange, "三级目录");
                    break;
                case 4:
                    row = RangeUtils.get_row_by_title(UsedRange, "四级目录");
                    break;
                case 5:
                    row = RangeUtils.get_row_by_title(UsedRange, "五级目录");
                    break;
                case 6:
                    row = RangeUtils.get_row_by_title(UsedRange, "六级目录");
                    break;
                case 7:
                    row = RangeUtils.get_row_by_title(UsedRange, "七级目录");
                    break;
                case 8:
                    row = RangeUtils.get_row_by_title(UsedRange, "八级目录");
                    break;
                case 9:
                    row = RangeUtils.get_row_by_title(UsedRange, "九级目录");
                    break;
                case 10:
                    row = RangeUtils.get_row_by_title(UsedRange, "十级目录");
                    break;
                default:
                    break;
            }
            if (row == 0) {
                return "";
            }
            String content = UsedRange[row, 2].Text;
            content = content.Trim();
            return content;
        }

        public String get_testStep_title() {
            int row = RangeUtils.get_row_by_title(UsedRange, "测试步骤");
            String content = UsedRange[row, 2].Text;
            content = content.Trim();
            return content;
        }

        public String get_testStep_index_title()
        {
            int row = RangeUtils.get_row_by_title(UsedRange, "测试步骤序号");
            String content = UsedRange[row, 2].Text;
            content = content.Trim();
            return content;
        }

        public String get_testcase_table_title()
        {
            int row = RangeUtils.get_row_by_title(UsedRange, "测试用例表题");
            String content = UsedRange[row, 2].Text;
            content = content.Trim();
            return content;
        }

        public String get_testcase_pass_title()
        {
            int row = RangeUtils.get_row_by_title(UsedRange, "测试用例执行结果");
            String content = UsedRange[row, 2].Text;
            content = content.Trim();
            return content;
        }

        public String get_testStep_pass_title()
        {
            int row = RangeUtils.get_row_by_title(UsedRange, "测试步骤通过与否");
            String content = UsedRange[row, 2].Text;
            content = content.Trim();
            return content;
        }

        public String get_testcase_relate_problem_title()
        {
            int row = RangeUtils.get_row_by_title(UsedRange, "关联的问题报告");
            String content = UsedRange[row, 2].Text;
            content = content.Trim();
            return content;
        }

        public String get_problem_relate_testcase_title()
        {
            int row = RangeUtils.get_row_by_title(UsedRange, "关联的测试用例");
            String content = UsedRange[row, 2].Text;
            content = content.Trim();
            return content;
        }

        public String get_problem_detail_title()
        {
            int row = RangeUtils.get_row_by_title(UsedRange, "问题详细描述");
            String content = UsedRange[row, 2].Text;
            content = content.Trim();
            return content;
        }

        public String get_problem_table_title()
        {
            int row = RangeUtils.get_row_by_title(UsedRange, "问题报告表题");
            String content = UsedRange[row, 2].Text;
            content = content.Trim();
            return content;
        }

        public String get_problem_type_title()
        {
            int row = RangeUtils.get_row_by_title(UsedRange, "问题类别");
            String content = UsedRange[row, 2].Text;
            content = content.Trim();
            return content;
        }

        public String get_problem_level_title()
        {            
            int row = RangeUtils.get_row_by_title(UsedRange, "问题级别");
            String content = UsedRange[row, 2].Text;
            content = content.Trim();
            return content;
        }

        public bool use_testcase_table_index()
        {            
            int row = RangeUtils.get_row_by_title(UsedRange, "测试用例表题编号");
            String content = UsedRange[row, 2].Text;
            content = content.Trim();
            if (!String.IsNullOrWhiteSpace(content) && content.Equals("是")) {
                return true;
            }
            return false;
        }

        public List<String> get_testor()
        {
            int row = RangeUtils.get_row_by_title(Cells, "测试人员");
            List<String> testors = new List<string>();
            for (int column = 2; column <= UsedRange.Columns.Count; column++)
            {
                String content = UsedRange[row, column].Text;
                content = Regex.Replace(content, @"[\r\n\s]*", "");
                if (!String.IsNullOrWhiteSpace(content)) {
                    testors.Add(content);
                }                
            }
            return testors;
        }

        public List<String> get_reportor()
        {
            int row = RangeUtils.get_row_by_title(Cells, "报告人");
            List<String> reportor = new List<string>();
            for (int column = 2; column <= UsedRange.Columns.Count; column++)
            {
                String content = UsedRange[row, column].Text;
                content = Regex.Replace(content, @"[\r\n\s]*", "");
                if (!String.IsNullOrWhiteSpace(content))
                {
                    reportor.Add(content);
                }
            }
            return reportor;
        }

        private void EditItemsSheet_Startup(object sender, System.EventArgs e)
        {
            hilight_and_lock_first_column();
        }

        private void EditItemsSheet_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO 设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.ActivateEvent += new Microsoft.Office.Interop.Excel.DocEvents_ActivateEventHandler(this.EditItemsSheet_ActivateEvent);
            this.Deactivate += new Microsoft.Office.Interop.Excel.DocEvents_DeactivateEventHandler(this.EditItemsSheet_Deactivate);
            this.Startup += new System.EventHandler(this.EditItemsSheet_Startup);
            this.Shutdown += new System.EventHandler(this.EditItemsSheet_Shutdown);

        }

        #endregion

        private void EditItemsSheet_ActivateEvent()
        {
            hilight_and_lock_first_column();
        }

        private void EditItemsSheet_Deactivate()
        {

        }
    }
}
