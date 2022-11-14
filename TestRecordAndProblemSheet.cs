using System;
using System.Collections.Generic;
using System.Data;
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
    public partial class TestRecordAndProblemSheet
    {
        public void update_relate_problem_and_testcase_from_testcase(String testcase, String problems)
        {
            int max_column = UsedRange.Columns.Count;
            int test_case_row = RangeUtils.get_row_by_title(Cells, testcase);
            if (test_case_row == 0) {
                int max_row = UsedRange.Rows.Count;
                Cells[max_row + 1, 1] = testcase;
                test_case_row = max_row + 1;
            }
            for (int column = 2; column <= max_column; column++) {
                if (problems.Contains(Cells[1,column].Text))
                {
                    Cells[test_case_row, column].Value = "是";
                    problems.Replace(Cells[1, column].Text, "");
                    problems.Replace("\n\n", "\n");
                }
                else {
                    Cells[test_case_row, column].Value = "";
                }
            }
        }
        public string get_problems_by_testcase_id(String testcase)
        {
            int max_column = UsedRange.Columns.Count;
            int test_case_row = RangeUtils.get_row_by_title(Cells, testcase);
            String ret = "";
            if (test_case_row == 0)
            {
                int max_row = UsedRange.Rows.Count;
                Cells[max_row + 1, 1] = testcase;
                test_case_row = max_row + 1;
            }

            for (int column = 2; column <= max_column; column++)
            {
                if (String.Equals(Cells[test_case_row, column].Value, "是")) {
                    if (String.IsNullOrEmpty(ret))
                    {
                        ret = Cells[1, column].Text;
                    }
                    else {
                        ret += "\n" + Cells[1, column].Text;
                    }
                }
            }
            return ret;
        }       

        public string get_testcases_by_problem_id(String problem_id)
        {
            int max_row = UsedRange.Rows.Count;
            int problem_column = RangeUtils.get_column_by_title(Cells, UsedRange.Columns.Count, problem_id);
            String ret = "";
            if (problem_column == 0)
            {
                int max_column = UsedRange.Columns.Count;
                Cells[1, max_column + 1] = problem_id;
                problem_column = max_column + 1;
            }

            for (int row = 2; row <= max_row; row++)
            {
                if (String.Equals(Cells[row, problem_column].Value, "是"))
                {
                    if (String.IsNullOrEmpty(ret))
                    {
                        ret = Cells[row,1].Text;
                    }
                    else
                    {
                        ret += "\n" + Cells[row,1].Text;
                    }
                }
            }
            return ret;
        }

        public void set_testcase_and_problem_cell(String testcase_id, String problem_id, String value)
        {
            int test_case_row = RangeUtils.get_row_by_title(Cells, testcase_id);
            int problem_column = RangeUtils.get_column_by_title(Cells, UsedRange.Columns.Count, problem_id);
            if (test_case_row == 0)
            {
                int max_row = UsedRange.Rows.Count;
                test_case_row = max_row + 1;
                Cells[test_case_row, 1].Value = testcase_id;
            }

            if (problem_column == 0)
            {
                int max_column = UsedRange.Columns.Count;
                problem_column = max_column + 1;
                Cells[1, problem_column].Value = problem_id;
            }

            Cells[test_case_row, problem_column].value = value;
        }

        public String get_testcase_and_problem_cell(String testcase_id, String problem_id)
        {
            int test_case_row = RangeUtils.get_row_by_title(Cells, testcase_id);
            int problem_column = RangeUtils.get_row_by_title(Cells, problem_id);

            if (test_case_row == 0 || problem_column == 0)
            {
                return "";
            }

            return Cells[test_case_row, problem_column].Text;
        }

        private void TestCaseAndProblemSheet_Startup(object sender, System.EventArgs e)
        {
        }

        private void TestCaseAndProblemSheet_Shutdown(object sender, System.EventArgs e)
        {
        }
        
        #region VSTO 设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(TestCaseAndProblemSheet_Startup);
            this.Shutdown += new System.EventHandler(TestCaseAndProblemSheet_Shutdown);
        }

        #endregion

    }
}
