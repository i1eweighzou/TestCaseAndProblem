using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Tools.Word;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using System.Text.RegularExpressions;
using System.IO;
using System.Diagnostics;
using Microsoft.Office.Interop.Word;
using System.Windows.Forms;
using Microsoft.Office.Tools;
using System.ComponentModel;

namespace TestCaseAndProblem
{
    public class OfficeWordUtils
    {
#region UTILS
        private static List<String> get_all_step(string string_steps)
        {
            List<String> ret = new List<String>();

            foreach (Match m in Regex.Matches(string_steps, @"^(.+)$", RegexOptions.Multiline))
            {
                String str = Regex.Replace(m.Value, @"^\d+\.\s*", "");
                ret.Add(Regex.Replace(str, @"[\r\n]*", ""));
            }
            return ret;
        }

        private static void replaceStr(Word.Range rng, ref object replaceKey, ref object replaceValue)
        {
            object replaceArea = Word.WdReplace.wdReplaceAll;
            Object Nothing = System.Reflection.Missing.Value;
            rng.Find.Execute(ref replaceKey, ref Nothing, ref Nothing, ref Nothing,
              ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing,
              ref replaceValue, ref replaceArea, ref Nothing, ref Nothing, ref Nothing,
              ref Nothing);
        }

        private static void replaceStrWithWild(Word.Range rng, ref object replaceKey, ref object replaceValue)
        {
            object replaceArea = Word.WdReplace.wdReplaceAll;
            Object Nothing = System.Reflection.Missing.Value;
            Object obj_true = true;
            rng.Find.Execute(ref replaceKey, ref Nothing, ref Nothing, ref obj_true,
              ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing,
              ref replaceValue, ref replaceArea, ref Nothing, ref Nothing, ref Nothing,
              ref Nothing);
        }

        private static Word.Row get_word_row_by_title(Word.Table table, String title, int column = 1)
        {

            foreach (Word.Row row in table.Rows)
            {
                String cell_str = Regex.Replace(row.Cells[column].Range.Text, @"[\r\n\a]*", "");
                if (String.Equals(cell_str, title))
                {
                    return row;
                }
            }
            return null;
        }

        private static int get_word_row_index_by_title(Word.Table table, String title, int column = 1) {
            Word.Row row = get_word_row_by_title(table, title, column);
            if (row != null) {
                return row.Index;
            }
            return 0;
        }

        private static Cell get_word_cell_by_title(Word.Table table, String title) {
            foreach (Cell cell in table.Range.Cells) {
                String cell_str = Regex.Replace(cell.Range.Text, @"[\r\n\a]*", "");
                if (String.Equals(cell_str, title))
                {
                    return cell;
                }
            }
            return null;
        }

        private static Dictionary<String, Cell> get_word_detail_edit_cells(Word.Table table, String titles)
        {
            Cell cell_start = null;
            Cell cell_end = null;
            Dictionary<String, Cell> dic = new Dictionary<String, Cell>();
            foreach (Cell cell in table.Range.Cells)
            {
                String cell_str = Regex.Replace(cell.Range.Text, @"[\r\n\a]*", "");
                Match match = Regex.Match(cell_str, @"(?<=\<)[^\>]+(?=\>)");
                if (match.Success) {
                    String title = match.Value;
                    if (titles.Contains(title)){
                        dic[title] = cell;
                        if (cell_start == null) {
                            cell_start = cell;
                        }
                        cell_end = cell;
                    }
                }
            }
            dic["cell_start"] = cell_start;
            dic["cell_end"] = cell_end;
            return dic;
        }

        private static void range_copy(Word.Range source, Word.Range target)
        {
            int num = 0;
            int limitNum = 5;
        retry:
            try
            {
                source.Copy();
                target.Paste();
            }
            catch (Exception)
            {
                num++;
                if (num > limitNum)
                {
                    throw;
                }
                goto retry;
            }
        }

        private static void range_copy(Word.Paragraph source, Word.Paragraph target)
        {
            range_copy(source.Range, target.Range);
        }

        private static Word.Cell cell_previous(Word.Cell cell, int count) {
            while (count-- > 0) {
                cell = cell.Previous;
            }
            return cell;
        }

        private static Word.Cell cell_next(Word.Cell cell, int count)
        {
            while (count-- > 0)
            {
                cell = cell.Next;
            }
            return cell;
        }
        private static String testRecord_p1_title;
        private static String testRecord_p2_title;
        private static String testRecord_p3_title;
        private static String testRecord_p4_title;
        private static String testRecord_p5_title;
        private static String testRecord_p6_title;
        private static String testRecord_p7_title;
        private static String testRecord_p8_title;
        private static String testRecord_p9_title;
        private static String testRecord_p10_title;

        public static string get_TestRecord_p1_title()
        {
            if (String.IsNullOrEmpty(testRecord_p1_title))
            {
                testRecord_p1_title = Globals.EditItemsSheet.get_outline_level_title(1);
            }
            return testRecord_p1_title;
        }

        public static string get_TestRecord_p2_title()
        {
            if (String.IsNullOrEmpty(testRecord_p2_title))
            {
                testRecord_p2_title = Globals.EditItemsSheet.get_outline_level_title(2);
            }
            return testRecord_p2_title;
        }

        public static string get_TestRecord_p3_title()
        {
            if (String.IsNullOrEmpty(testRecord_p3_title))
            {
                testRecord_p3_title = Globals.EditItemsSheet.get_outline_level_title(3);
            }
            return testRecord_p3_title;
        }

        public static string get_TestRecord_p4_title()
        {
            if (String.IsNullOrEmpty(testRecord_p4_title))
            {
                testRecord_p4_title = Globals.EditItemsSheet.get_outline_level_title(4);
            }
            return testRecord_p4_title;
        }

        public static string get_TestRecord_p5_title()
        {
            if (String.IsNullOrEmpty(testRecord_p5_title))
            {
                testRecord_p5_title = Globals.EditItemsSheet.get_outline_level_title(5);
            }
            return testRecord_p5_title;
        }

        public static string get_TestRecord_p6_title()
        {
            if (String.IsNullOrEmpty(testRecord_p6_title))
            {
                testRecord_p6_title = Globals.EditItemsSheet.get_outline_level_title(6);
            }
            return testRecord_p6_title;
        }

        public static string get_TestRecord_p7_title()
        {
            if (String.IsNullOrEmpty(testRecord_p7_title))
            {
                testRecord_p7_title = Globals.EditItemsSheet.get_outline_level_title(7);
            }
            return testRecord_p7_title;
        }

        public static string get_TestRecord_p8_title()
        {
            if (String.IsNullOrEmpty(testRecord_p8_title))
            {
                testRecord_p8_title = Globals.EditItemsSheet.get_outline_level_title(8);
            }
            return testRecord_p8_title;
        }

        public static string get_TestRecord_p9_title()
        {
            if (String.IsNullOrEmpty(testRecord_p9_title))
            {
                testRecord_p9_title = Globals.EditItemsSheet.get_outline_level_title(9);
            }
            return testRecord_p9_title;
        }

        public static string get_TestRecord_p10_title()
        {
            if (String.IsNullOrEmpty(testRecord_p10_title))
            {
                testRecord_p10_title = Globals.EditItemsSheet.get_outline_level_title(10);
            }
            return testRecord_p10_title;
        }
        private static int get_paragraph_level_column(Excel.Range cells, int level)
        {
            String level_str;
            switch (level)
            {
                case 1:
                    level_str = get_TestRecord_p1_title();
                    break;
                case 2:
                    level_str = get_TestRecord_p2_title();
                    break;
                case 3:
                    level_str = get_TestRecord_p3_title();
                    break;
                case 4:
                    level_str = get_TestRecord_p4_title();
                    break;
                case 5:
                    level_str = get_TestRecord_p5_title();
                    break;
                case 6:
                    level_str = get_TestRecord_p6_title();
                    break;
                case 7:
                    level_str = get_TestRecord_p7_title();
                    break;
                case 8:
                    level_str = get_TestRecord_p8_title();
                    break;
                case 9:
                    level_str = get_TestRecord_p9_title();
                    break;
                case 10:
                    level_str = get_TestRecord_p10_title();
                    break;
                default:
                    return 0;
            }
            return RangeUtils.get_column_by_title(cells, level_str);
        }
        #endregion

        #region TESTRECORD
        private static void check_pass(Excel.Range cells, ref string warnningText, int column_testcase_pass, int row, int column)
        {
            String steps_pass = cells[row, column].Text;
            String testCase_pass = cells[row, column_testcase_pass].Text;
            testCase_pass = testCase_pass.Replace("\r", "").Replace("\n", "").Replace("\a", "");
            if (String.IsNullOrWhiteSpace(testCase_pass)
                || steps_pass.Contains("n") && String.Equals(testCase_pass, "通过")
                || !steps_pass.Contains("y") && String.Equals(testCase_pass, "不通过"))
            {
                warnningText += row + "行不匹配\r\n";
            }
        }

        private static String testStep_title;
        private static String get_testStep_title() {
            if (String.IsNullOrWhiteSpace(testStep_title))
            {
                testStep_title = Globals.EditItemsSheet.get_testStep_title();
            }
            return testStep_title;
        }
        private static int get_testStep_info_row(Word.Document doc_template)
        {
            String testStep_title = get_testStep_title();
            int detailEditRow = get_word_row_index_by_title(doc_template.Tables[1], testStep_title) + 2;
            if (detailEditRow == 0)
            {
                MessageBox.Show("模板中必须包含测试步骤，测试步骤下一行为标题行，再下一行为空白行");
            }

            return detailEditRow;
        }        

        private static bool add_paragraph_level(Excel.Range cells, Word.Document doc, int level, Paragraph p, int p_column, ref string last_p_text, int row)
        {
            object replaceKey;
            object replaceValue;
            bool ret = false;
            string replace_str;

            switch (level) {
                case 1:
                    replace_str = get_TestRecord_p1_title();
                    break;
                case 2:
                    replace_str = get_TestRecord_p2_title();
                    break;
                case 3:
                    replace_str = get_TestRecord_p3_title();
                    break;
                case 4:
                    replace_str = get_TestRecord_p4_title();
                    break;
                case 5:
                    replace_str = get_TestRecord_p5_title();
                    break;
                case 6:
                    replace_str = get_TestRecord_p6_title();
                    break;
                case 7:
                    replace_str = get_TestRecord_p7_title();
                    break;
                case 8:
                    replace_str = get_TestRecord_p8_title();
                    break;
                case 9:
                    replace_str = get_TestRecord_p9_title();
                    break;
                case 10:
                    replace_str = get_TestRecord_p10_title();
                    break;
                default:
                    return ret;
            }

            if (p != null && p_column != 0)
            {
                String value = cells[row, p_column].Text;
                String value_title = cells[1, p_column].Text;
                value = value.Trim();

                if (String.Equals(value, "Skip") || String.Equals(value_title, "用例名称"))
                {
                    int testcase_name_column = get_usecase_name_column(cells);
                    if (testcase_name_column == 0)
                    {
                        throw (new Exception("必须包含用例名称列"));
                    }
                    value = cells[row, testcase_name_column].Text;
                    ret = true;
                }


                if (!String.IsNullOrEmpty(value) && !value.Equals(last_p_text))
                {
                    Paragraph paragraph = doc.Paragraphs.Add();
                    range_copy(p, paragraph);
                    replaceKey = "<" + replace_str + ">";
                    replaceValue = value;
                    paragraph = paragraph.Previous(1);
                    replaceStr(paragraph.Range, ref replaceKey, ref replaceValue);
                    last_p_text = value;
                }
            }

            return ret;
        }

        private static int usecase_name_column = 0;
        private static int get_usecase_name_column(Excel.Range cells)
        {
            if (usecase_name_column == 0) {
                usecase_name_column = RangeUtils.get_column_by_title(cells, "用例名称");
            }
            return usecase_name_column;
        }

        private static void add_testcase_table_title(Excel.Range cells, Word.Document doc, Paragraph p, int row, bool use_table_index)
        {
            int testcase_name_column = get_usecase_name_column(cells);
            if (testcase_name_column == 0)
            {
                throw (new Exception("必须包含用例名称列"));
            }
            object replaceKey;
            object replaceValue;

            string replace_str = get_testcase_table_title();

            if (p != null && testcase_name_column != 0)
            {
                String value = cells[row, testcase_name_column].Text;
                Paragraph paragraph = doc.Paragraphs.Add();
                range_copy(p, paragraph);
                replaceKey = "<" + replace_str + ">";
                if (use_table_index)
                {
                    replaceValue = "表" + (row - 1) + " " + value;
                }
                else
                {
                    replaceValue = value;
                }

                paragraph = paragraph.Previous(1);
                replaceStr(paragraph.Range, ref replaceKey, ref replaceValue);
            }
        }

        private static void add_testcase_outline_level(Excel.Range cells, Word.Document doc, Paragraph p, int row)
        {
            int testcase_name_column = get_usecase_name_column(cells);
            if (testcase_name_column == 0)
            {
                throw (new Exception("必须包含用例名称列"));
            }
            object replaceKey;
            object replaceValue;
            string replace_str = get_testcase_table_title();

            if (p != null && testcase_name_column != 0)
            {
                String value = cells[row, testcase_name_column].Text;
                Paragraph paragraph = doc.Paragraphs.Add();
                range_copy(p, paragraph);
                replaceKey = "<" + replace_str + ">";
                replaceValue = value;
                paragraph = paragraph.Previous(1);
                replaceStr(paragraph.Range, ref replaceKey, ref replaceValue);
            }

            
        }

        private static string testcase_table_title;
        private static string get_testcase_table_title()
        {
            if (String.IsNullOrWhiteSpace(testcase_table_title)) {
                testcase_table_title = Globals.EditItemsSheet.get_testcase_table_title();
            }
            return testcase_table_title;
        }

        private static int get_testcase_pass_column(Excel.Range cells)
        {
            string testCase_pass_title = get_testcase_pass_title();
            int column_testcase_pass = get_column_testcase_pass(cells, testCase_pass_title);
            return column_testcase_pass;
        }

        private static int column_testcase_pass;
        private static int get_column_testcase_pass(Excel.Range cells, string testCase_pass_title)
        {
            if (column_testcase_pass == 0) {
                column_testcase_pass = RangeUtils.get_column_by_title(cells, testCase_pass_title);
            }
            return column_testcase_pass;
        }

        private static string testCase_pass_title;
        private static string get_testcase_pass_title()
        {
            if (String.IsNullOrWhiteSpace(testCase_pass_title)) {
                testCase_pass_title = Globals.EditItemsSheet.get_testcase_pass_title();
            }
            return testCase_pass_title;
        }

        private static void add_level_pragraphs(Excel.Range cells, Word.Document doc, Paragraph p1, Paragraph p2, Paragraph p3, Paragraph p4, Paragraph p5, Paragraph p6, Paragraph p7, Paragraph p8, ref string last_p1_text, ref string last_p2_text, ref string last_p3_text, ref string last_p4_text, ref string last_p5_text, ref string last_p6_text, ref string last_p7_text, ref string last_p8_text, int p1_column, int p2_column, int p3_column, int p4_column, int p5_column, int p6_column, int p7_column, int p8_column, int row)
        {
            if (add_paragraph_level(cells, doc, 1, p1, p1_column, ref last_p1_text, row)) {
                return;
            }
            if (add_paragraph_level(cells, doc, 2, p2, p2_column, ref last_p2_text, row)) {
                return;
            }
            if (add_paragraph_level(cells, doc, 3, p3, p3_column, ref last_p3_text, row)) {
                return;
            }
            if (add_paragraph_level(cells, doc, 4, p4, p4_column, ref last_p4_text, row)) {
                return;
            }
            if (add_paragraph_level(cells, doc, 5, p5, p5_column, ref last_p5_text, row)) {
                return;
            }
            if (add_paragraph_level(cells, doc, 6, p6, p6_column, ref last_p6_text, row)) {
                return;
            }
            if (add_paragraph_level(cells, doc, 7, p7, p7_column, ref last_p7_text, row)) {
                return;
            }
            if (add_paragraph_level(cells, doc, 8, p8, p8_column, ref last_p8_text, row)) {
                return;
            }
        }
        #endregion
        
        #region PROBLEM
        private static Paragraph get_problem_table_title_paragraph(Word.Document doc_template)
        {
            Word.Paragraph p1 = null;
            String problem_table_title = Globals.EditItemsSheet.get_problem_table_title();
            foreach (Word.Paragraph p in doc_template.Paragraphs)
            {
                String pstr = p.Range.Text;
                if (pstr.Contains("<" + problem_table_title + ">"))
                {
                    p1 = p;
                    break;
                }
            }

            return p1;
        }

        private static void add_problem_table_title(Word.Document doc, Paragraph p_table_title)
        {
            object replaceKey;
            object replaceValue;
            String problem_table_title = Globals.EditItemsSheet.get_problem_table_title();
            if (p_table_title != null)
            {
                Paragraph paragraph_1;
                paragraph_1 = doc.Paragraphs.Add();
                range_copy(p_table_title, paragraph_1);
                replaceKey = "<" + problem_table_title + ">";
                replaceValue = "问题报告单";
                paragraph_1 = paragraph_1.Previous(1);
                replaceStr(paragraph_1.Range, ref replaceKey, ref replaceValue);
            }
        }

        private static void set_problem_type(Excel.Range cells, int row, Table table, int column)
        {
            object replaceKey;
            object replaceValue;
            switch (cells[row, column].Text)
            {
                case "程序":
                    replaceKey = "<程序>";
                    replaceValue = "■";
                    replaceStr(table.Range, ref replaceKey, ref replaceValue);
                    replaceKey = "<文档>";
                    replaceValue = "□";
                    replaceStr(table.Range, ref replaceKey, ref replaceValue);
                    replaceKey = "<设计>";
                    replaceValue = "□";
                    replaceStr(table.Range, ref replaceKey, ref replaceValue);
                    replaceKey = "<其它>";
                    replaceValue = "□";
                    replaceStr(table.Range, ref replaceKey, ref replaceValue);
                    break;
                case "文档":
                    replaceKey = "<程序>";
                    replaceValue = "□";
                    replaceStr(table.Range, ref replaceKey, ref replaceValue);
                    replaceKey = "<文档>";
                    replaceValue = "■";
                    replaceStr(table.Range, ref replaceKey, ref replaceValue);
                    replaceKey = "<设计>";
                    replaceValue = "□";
                    replaceStr(table.Range, ref replaceKey, ref replaceValue);
                    replaceKey = "<其它>";
                    replaceValue = "□";
                    replaceStr(table.Range, ref replaceKey, ref replaceValue);
                    break;
                case "设计":
                    replaceKey = "<程序>";
                    replaceValue = "□";
                    replaceStr(table.Range, ref replaceKey, ref replaceValue);
                    replaceKey = "<文档>";
                    replaceValue = "□";
                    replaceStr(table.Range, ref replaceKey, ref replaceValue);
                    replaceKey = "<设计>";
                    replaceValue = "■";
                    replaceStr(table.Range, ref replaceKey, ref replaceValue);
                    replaceKey = "<其它>";
                    replaceValue = "□";
                    replaceStr(table.Range, ref replaceKey, ref replaceValue);
                    break;
                case "其它":
                    replaceKey = "<程序>";
                    replaceValue = "□";
                    replaceStr(table.Range, ref replaceKey, ref replaceValue);
                    replaceKey = "<文档>";
                    replaceValue = "□";
                    replaceStr(table.Range, ref replaceKey, ref replaceValue);
                    replaceKey = "<设计>";
                    replaceValue = "□";
                    replaceStr(table.Range, ref replaceKey, ref replaceValue);
                    replaceKey = "<其它>";
                    replaceValue = "■";
                    replaceStr(table.Range, ref replaceKey, ref replaceValue);
                    break;
                default:
                    break;
            }
        }

        private static void set_problem_level(Excel.Range cells, int row, Table table, int column)
        {
            object replaceKey;
            object replaceValue;
            switch (cells[row, column].Text)
            {
                case "致命":
                    replaceKey = "<致命>";
                    replaceValue = "■";
                    replaceStr(table.Range, ref replaceKey, ref replaceValue);
                    replaceKey = "<严重>";
                    replaceValue = "□";
                    replaceStr(table.Range, ref replaceKey, ref replaceValue);
                    replaceKey = "<一般>";
                    replaceValue = "□";
                    replaceStr(table.Range, ref replaceKey, ref replaceValue);
                    replaceKey = "<轻微>";
                    replaceValue = "□";
                    replaceStr(table.Range, ref replaceKey, ref replaceValue);
                    break;
                case "严重":
                    replaceKey = "<致命>";
                    replaceValue = "□";
                    replaceStr(table.Range, ref replaceKey, ref replaceValue);
                    replaceKey = "<严重>";
                    replaceValue = "■";
                    replaceStr(table.Range, ref replaceKey, ref replaceValue);
                    replaceKey = "<一般>";
                    replaceValue = "□";
                    replaceStr(table.Range, ref replaceKey, ref replaceValue);
                    replaceKey = "<轻微>";
                    replaceValue = "□";
                    replaceStr(table.Range, ref replaceKey, ref replaceValue);
                    break;
                case "一般":
                    replaceKey = "<致命>";
                    replaceValue = "□";
                    replaceStr(table.Range, ref replaceKey, ref replaceValue);
                    replaceKey = "<严重>";
                    replaceValue = "□";
                    replaceStr(table.Range, ref replaceKey, ref replaceValue);
                    replaceKey = "<一般>";
                    replaceValue = "■";
                    replaceStr(table.Range, ref replaceKey, ref replaceValue);
                    replaceKey = "<轻微>";
                    replaceValue = "□";
                    replaceStr(table.Range, ref replaceKey, ref replaceValue);
                    break;
                case "轻微":
                    replaceKey = "<致命>";
                    replaceValue = "□";
                    replaceStr(table.Range, ref replaceKey, ref replaceValue);
                    replaceKey = "<严重>";
                    replaceValue = "□";
                    replaceStr(table.Range, ref replaceKey, ref replaceValue);
                    replaceKey = "<一般>";
                    replaceValue = "□";
                    replaceStr(table.Range, ref replaceKey, ref replaceValue);
                    replaceKey = "<轻微>";
                    replaceValue = "■";
                    replaceStr(table.Range, ref replaceKey, ref replaceValue);
                    break;
                default:
                    break;
            }
        }

        private static void set_relate_testcase(Excel.Range cells, Word.Document doc, int row, Table table)
        {
            String detail_edit_item_problem_template = Globals.EditItemsSheet.get_detail_edit_item_problem_template();
            Dictionary<String, Cell> word_detail_edit_cells =  get_word_detail_edit_cells(table, detail_edit_item_problem_template);
            Word.Cell cell_start = word_detail_edit_cells["cell_start"];
            Word.Cell cell_end = word_detail_edit_cells["cell_end"];

            int step = 1;
            int column = RangeUtils.get_column_by_title(cells, "问题标识");
            String problem_id = cells[row, column].Text;
            String testcases = Globals.TestRecordAndProblemSheet.get_testcases_by_problem_id(problem_id);
            MatchCollection matchCollection = Regex.Matches(testcases, @"^(.+)$", RegexOptions.Multiline);
            Cell new_cell_start;
            foreach (Match m in matchCollection)
            {
                if (String.IsNullOrWhiteSpace(m.Value))
                {
                    continue;
                }
                string testcase_id = m.Value;
                string testcase_name = Globals.TestRecordSheet.get_testcase_name_by_id(testcase_id);

                if (step != matchCollection.Count)
                {
                    object start = cell_start.Range.Start;
                    object end = cell_end.Range.End;
                    Word.Range range = doc.Range(ref start, ref end);
                    range.Rows.Add(range);
                    new_cell_start = cell_previous(cell_start, word_detail_edit_cells.Count - 2);
                }
                else {
                    new_cell_start = cell_start;
                }
                foreach (KeyValuePair<String, Cell> kp in word_detail_edit_cells)
                {
                    switch (kp.Key)
                    {
                        case "测试用例标识":
                            new_cell_start.Range.Text = testcase_id;
                            new_cell_start = new_cell_start.Next;
                            break;
                        case "测试用例名称":
                            new_cell_start.Range.Text = testcase_name;
                            new_cell_start = new_cell_start.Next;
                            break;
                        default:
                            break;
                    }                    
                }
                step++;
            }
        }

        #endregion
        private static int get_detailEditRow(Word.Document doc_template)
        {
            String problem_detail_title = getProblem_detail_title();
            int detailEditRow = get_word_row_index_by_title(doc_template.Tables[1], problem_detail_title);
            if (detailEditRow == 0)
            {
                throw new Exception("模板中必须包含详细描述，详细描述下一行为空白行");
            }

            return detailEditRow;
        }
        private static string problem_detail_title;
        private static string getProblem_detail_title()
        {
            if (String.IsNullOrWhiteSpace(problem_detail_title)) {
                problem_detail_title = Globals.EditItemsSheet.get_problem_detail_title();
            }
            return problem_detail_title;
        }

        private static Word.Cell get_detailEditCell(Word.Table table)
        {
            String problem_detail_title = getProblem_detail_title();
            Word.Cell cell = get_word_cell_by_title(table, problem_detail_title);
            if (cell == null)
            {
                throw new Exception("模板中必须包含详细描述，详细描述下一行为空白行");
            }

            return cell.Next;
        }

        private static Word.Cell get_relate_testcase_Cell(Word.Table table)
        {
            string problem_relate_testcase_title = get_problem_relate_testcase_title();
            Word.Cell cell = get_word_cell_by_title(table, problem_relate_testcase_title);
            if (cell == null)
            {
                throw new Exception("模板中必须包含关联的测试用例，详细描述下一行为空白行");
            }

            return cell.Next.Next.Next;
        }

        private static string problem_relate_testcase_title;
        private static string get_problem_relate_testcase_title()
        {
            if (String.IsNullOrWhiteSpace(problem_relate_testcase_title)) {
                problem_relate_testcase_title = Globals.EditItemsSheet.get_problem_relate_testcase_title();
            }
            return problem_relate_testcase_title;
        }

        private static void set_page_box_outline_by_template(Word.Document doc, Word.Document doc_template)
        {
            doc.PageSetup.TopMargin = doc_template.PageSetup.TopMargin;
            doc.PageSetup.BottomMargin = doc_template.PageSetup.BottomMargin;
            doc.PageSetup.LeftMargin = doc_template.PageSetup.LeftMargin;
            doc.PageSetup.RightMargin = doc_template.PageSetup.RightMargin;
            doc.PageSetup.Gutter = doc_template.PageSetup.Gutter;
        }
        private static void remove_replace_with_white_space(out object replaceKey, out object replaceValue, Table table)
        {
            replaceKey = @"\<?@\>";
            replaceValue = "";
            replaceStrWithWild(table.Range, ref replaceKey, ref replaceValue);
        }

        private static Table copy_row_table(Word.Document doc, Word.Document doc_template, int row)
        {
            Paragraph paragraph_table = doc.Paragraphs.Add();
            range_copy(doc_template.Tables[1].Range.Duplicate, paragraph_table.Range);
            Table table = doc.Tables[row - 1];
            return table;
        }
        private static void get_paragraph_outline_template(Word.Document doc_template, ref Paragraph p1, ref Paragraph p2, ref Paragraph p3, ref Paragraph p4, ref Paragraph p5, ref Paragraph p6, ref Paragraph p7, ref Paragraph p8, ref Paragraph p_table_title)
        {
            foreach (Word.Paragraph p in doc_template.Paragraphs)
            {
                String pstr = p.Range.Text;
                if (p1 == null && pstr.Contains("<" + get_TestRecord_p1_title() + ">"))
                {
                    p1 = p;
                }
                else if (p2 == null && pstr.Contains("<" + get_TestRecord_p2_title() + ">"))
                {
                    p2 = p;
                }
                else if (p3 == null && pstr.Contains("<" + get_TestRecord_p3_title() + ">"))
                {
                    p3 = p;
                }
                else if (p4 == null && pstr.Contains("<" + get_TestRecord_p4_title() + ">"))
                {
                    p4 = p;
                }
                else if (p5 == null && pstr.Contains("<" + get_TestRecord_p5_title() + ">"))
                {
                    p5 = p;
                }
                else if (p6 == null && pstr.Contains("<" + get_TestRecord_p6_title() + ">"))
                {
                    p6 = p;
                }
                else if (p7 == null && pstr.Contains("<" + get_TestRecord_p7_title() + ">"))
                {
                    p7 = p;
                }
                else if (p8 == null && pstr.Contains("<" + get_TestRecord_p8_title() + ">"))
                {
                    p8 = p;
                }
                else if (p_table_title == null &&pstr.Contains("<" + get_testcase_table_title() + ">"))
                {
                    p_table_title = p;
                }
            }
        }
        public static void openWordFile(string filepath)
        {
            object miss = System.Reflection.Missing.Value;

            Word.Application appWord = null;
            Word.Document doc = null;

            try
            {
                appWord = new Word.Application
                {
                    Visible = true
                };

                object objTrue = true;
                object objFalse = false;
                object objFilePath = filepath;
                object objDocType = WdDocumentType.wdTypeDocument;
                doc = appWord.Documents.Add(ref objFilePath, ref objFalse, ref objDocType, ref objTrue);            
            }
            catch (System.Exception ex)
            {
                Debug.WriteLine(ex.Message);
                MessageBox.Show(ex.Message);
            }
        }

        public static void createWordTestRecordByDocmentTemplate(string temppath, string exportFilefullname, Excel.Range cells, BackgroundWorker backgroundWorker)
        {
            object miss = System.Reflection.Missing.Value;
            object outputfullfilename = exportFilefullname;

            Word.Application appWord = null;
            Word.Document doc = null;
            Word.Document doc_template = null;

            try
            {
                appWord = new Word.Application
                {
                    Visible = false
                };

                object objTrue = true;
                object objFalse = false;
                object objTemplate = temppath;
                object objDocType = WdDocumentType.wdTypeDocument;
                object replaceKey;
                object replaceValue;
                doc_template = appWord.Documents.Add(ref objTemplate, ref objFalse, ref objDocType, ref objTrue);
                doc = appWord.Documents.Add();
                backgroundWorker.ReportProgress(0, "开始处理");
                bool use_table_index = Globals.EditItemsSheet.use_testcase_table_index();

                set_page_box_outline_by_template(doc, doc_template);

                String detailEditItems = Globals.EditItemsSheet.get_detail_edit_item_testcase();
                String testStep_index_title = Globals.EditItemsSheet.get_testStep_index_title();
                int detailEditRow = get_testStep_info_row(doc_template);
                String testStep_pass_title = Globals.EditItemsSheet.get_testStep_pass_title();
                String warnningText = "";

                Word.Paragraph p1 = null;
                Word.Paragraph p2 = null;
                Word.Paragraph p3 = null;
                Word.Paragraph p4 = null;
                Word.Paragraph p5 = null;
                Word.Paragraph p6 = null;
                Word.Paragraph p7 = null;
                Word.Paragraph p8 = null;
                Word.Paragraph p_table_title = null;

                String last_p1_text = "";
                String last_p2_text = "";
                String last_p3_text = "";
                String last_p4_text = "";
                String last_p5_text = "";
                String last_p6_text = "";
                String last_p7_text = "";
                String last_p8_text = "";

                int p1_column = get_paragraph_level_column(cells, 1);
                int p2_column = get_paragraph_level_column(cells, 2);
                int p3_column = get_paragraph_level_column(cells, 3);
                int p4_column = get_paragraph_level_column(cells, 4);
                int p5_column = get_paragraph_level_column(cells, 5);
                int p6_column = get_paragraph_level_column(cells, 6);
                int p7_column = get_paragraph_level_column(cells, 7);
                int p8_column = get_paragraph_level_column(cells, 8);
                int column_testcase_pass = get_testcase_pass_column(cells);
                int last_percent = 0;
                get_paragraph_outline_template(doc_template, ref p1, ref p2, ref p3, ref p4, ref p5, ref p6, ref p7, ref p8, ref p_table_title);

                int cell_max_row = RangeUtils.get_max_row(cells);
                int cell_max_column = RangeUtils.get_max_column(cells);
                for (int row = 2; row <= cell_max_row; row++)
                {
                    if (backgroundWorker.CancellationPending) {
                        break;
                    }
                    add_level_pragraphs(cells, doc, p1, p2, p3, p4, p5, p6, p7, p8, ref last_p1_text, ref last_p2_text, ref last_p3_text, ref last_p4_text, ref last_p5_text, ref last_p6_text, ref last_p7_text, ref last_p8_text, p1_column, p2_column, p3_column, p4_column, p5_column, p6_column, p7_column, p8_column, row);

                    add_testcase_table_title(cells, doc, p_table_title, row, use_table_index);

                    Table table = copy_row_table(doc, doc_template, row);
                    int max_step_len = 0;
                    Dictionary<String, List<String>> detailEditItemsDict = new Dictionary<String, List<String>>();
                    for (int column = 1; column <= cell_max_column; column++)
                    {
                        String title = cells[1, column].Text;
                        if (detailEditItems.Contains(title))
                        {
                            if (!detailEditItemsDict.ContainsKey(title))
                            {
                                List<String> steps = get_all_step(cells[row, column].Text);
                                if (steps.Count > max_step_len)
                                {
                                    max_step_len = steps.Count;
                                }
                                detailEditItemsDict.Add(title, steps);
                            }
                        }
                        else
                        {
                            replaceKey = "<" + title + ">";
                            replaceValue = cells[row, column].Text;
                            replaceStr(table.Range, ref replaceKey, ref replaceValue);
                        }
                    }
                    string detailEditItemsTemplate = get_detailEditItemsTemplate();
                    Dictionary<String, Cell> word_detail_edit_cells = get_word_detail_edit_cells(table, detailEditItemsTemplate);
                    Word.Cell cell_start = word_detail_edit_cells["cell_start"];
                    Word.Cell cell_end = word_detail_edit_cells["cell_end"];
                    Cell new_cell_start;
                    for (int si = 0; si < max_step_len; si++)
                    {
                        if (si != max_step_len - 1)
                        {
                            object start = cell_start.Range.Start;
                            object end = cell_end.Range.End;
                            Word.Range range = doc.Range(ref start, ref end);
                            range.Rows.Add(range);
                            new_cell_start = cell_previous(cell_start, word_detail_edit_cells.Count - 2);
                        }
                        else
                        {
                            new_cell_start = cell_start;
                        }

                        foreach (KeyValuePair<String, Cell> kp in word_detail_edit_cells)
                        {
                            String title = kp.Key;
                            if (String.IsNullOrEmpty(title))
                            {
                                continue;
                            }
                            else if (String.Equals(testStep_index_title, title))
                            {
                                new_cell_start.Range.Text = (si + 1).ToString();
                                new_cell_start = new_cell_start.Next;
                                continue;
                            }
                            else if (detailEditItemsDict.ContainsKey(title))
                            {
                                if (detailEditItemsDict[title].Count > si)
                                {
                                    String str_content = detailEditItemsDict[title].ElementAt<String>(si);
                                    if (title.Contains("图片"))
                                    {
                                        if (!String.IsNullOrEmpty(str_content))
                                        {
                                            String file_path = PictureUtils.get_picture_path(str_content);
                                            new_cell_start.Range.InlineShapes.AddPicture(file_path);
                                            new_cell_start = new_cell_start.Next;
                                        }
                                    }
                                    else
                                    {
                                        new_cell_start.Range.Text = str_content;
                                        new_cell_start = new_cell_start.Next;
                                    }
                                }
                            }
                        }
                    }

                    remove_replace_with_white_space(out replaceKey, out replaceValue, table);
                    if ((last_percent) < ((row - 1) * 100) / cell_max_row)
                    {
                        last_percent = ((row - 1) * 100) / cell_max_row;
                        backgroundWorker.ReportProgress(last_percent, "处理第" + row + "行");
                    }
                }
                object fileformat = WdSaveFormat.wdFormatDocument;
                doc.SaveAs2(outputfullfilename);

                if (!String.IsNullOrEmpty(warnningText))
                {
                    MessageBox.Show(warnningText, "不匹配警告");
                }
            }
            finally
            {
                object doNotSaveChanges = WdSaveOptions.wdDoNotSaveChanges;
                doc?.Close(ref doNotSaveChanges, ref miss, ref miss);
                doc_template?.Close(ref doNotSaveChanges, ref miss, ref miss);
                appWord.Application.Quit(ref miss, ref miss, ref miss);
                doc = null;
                doc_template = null;
                appWord = null;
            }
        }

        private static string detailEditItemsTemplate;
        private static string get_detailEditItemsTemplate()
        {
            if (String.IsNullOrWhiteSpace(detailEditItemsTemplate)) {
                detailEditItemsTemplate = Globals.EditItemsSheet.get_detail_edit_item_testcase_template();
            }
            return detailEditItemsTemplate;
        }

        public static void createWordProblemByDocmentTemplate(string temppath, string exportFilefullname, Excel.Range cells, BackgroundWorker backgroundWorker)
        {
            object miss = System.Reflection.Missing.Value;
            object outputfullfilename = exportFilefullname;

            Word.Application appWord = null;
            Word.Document doc = null;
            Word.Document doc_template = null;
            int last_percent = 0;

            try
            {
                appWord = new Word.Application
                {
                    Visible = false
                };

                object objTrue = true;
                object objFalse = false;
                object objTemplate = temppath;
                object objDocType = WdDocumentType.wdTypeDocument;
                object replaceKey;
                object replaceValue;
                doc_template = appWord.Documents.Add(ref objTemplate, ref objFalse, ref objDocType, ref objTrue);
                doc = appWord.Documents.Add();
                backgroundWorker.ReportProgress(0, "开始处理");
                set_page_box_outline_by_template(doc, doc_template);
                String detailEditItems = Globals.EditItemsSheet.get_detail_edit_item_problem();
                String relate_testcase_title = Globals.EditItemsSheet.get_problem_relate_testcase_title();
                String problem_type_title = Globals.EditItemsSheet.get_problem_type_title();
                String problem_level_title = Globals.EditItemsSheet.get_problem_level_title();
                Paragraph p_table_title = get_problem_table_title_paragraph(doc_template);

                int cell_max_row = RangeUtils.get_max_row(cells);
                int cell_max_column = RangeUtils.get_max_column(cells);
                for (int row = 2; row <= cell_max_row; row++)
                {
                    if (backgroundWorker.CancellationPending)
                    {
                        break;
                    }
                    add_problem_table_title(doc, p_table_title);

                    Table table = copy_row_table(doc, doc_template, row);
                    int max_step_len = 0;
                    Dictionary<String, List<String>> detailEditItemsDict = new Dictionary<String, List<String>>();
                    for (int column = 1; column <= cell_max_column; column++)
                    {
                        String title = cells[1, column].Text;
                        if (detailEditItems.Contains(title))
                        {
                            if (!detailEditItemsDict.ContainsKey(title))
                            {
                                List<String> steps = get_all_step(cells[row, column].Text);
                                if (steps.Count > max_step_len)
                                {
                                    max_step_len = steps.Count;
                                }
                                detailEditItemsDict.Add(title, steps);
                            }
                        }
                        else if (String.Equals(problem_type_title, title))
                        {
                            set_problem_type(cells, row, table, column);

                        }
                        else if (String.Equals(problem_level_title, title))
                        {
                            set_problem_level(cells, row, table, column);
                        }
                        else
                        {

                            replaceKey = "<" + title + ">";
                            replaceValue = cells[row, column].Text;
                            replaceStr(table.Range, ref replaceKey, ref replaceValue);
                        }
                    }
                    set_relate_testcase(cells, doc, row, table);                    

                    for (int si = 0; si < max_step_len; si++)
                    {
                        foreach (KeyValuePair<string, List<string>> ergodic in detailEditItemsDict)
                        {
                            if (ergodic.Value.Count > si)
                            {
                                String str_content = ergodic.Value.ElementAt<String>(si);
                                Word.Cell detailCell = get_detailEditCell(table);
                                Word.Range wt_rng = detailCell.Range;
                                if (ergodic.Key.Contains("图片"))
                                {
                                    Word.Paragraph paragraph_table_cell;
                                    if (!String.IsNullOrEmpty(str_content))
                                    {
                                        paragraph_table_cell = wt_rng.Paragraphs.Add();
                                        object style = "无间隔";
                                        paragraph_table_cell.Range.set_Style(WdBuiltinStyle.wdStyleBodyText);
                                        String file_path = PictureUtils.get_picture_path(str_content);
                                        paragraph_table_cell.Range.InlineShapes.AddPicture(file_path);
                                    }
                                }
                                else if (ergodic.Key.Contains("详细描述"))
                                {
                                    Word.Paragraph paragraph_table_cell;
                                    if (si == 0)
                                    {
                                        paragraph_table_cell = wt_rng.Paragraphs[1];
                                    }
                                    else
                                    {
                                        paragraph_table_cell = wt_rng.Paragraphs.Add();
                                    }
                                    paragraph_table_cell.Range.Text = (si + 1) + "." + str_content;
                                }
                                else
                                {
                                    Word.Paragraph paragraph_table_cell;
                                    if (si == 0)
                                    {
                                        paragraph_table_cell = wt_rng.Paragraphs[1];
                                    }
                                    else
                                    {
                                        paragraph_table_cell = wt_rng.Paragraphs.Add();
                                    }
                                    paragraph_table_cell.Range.Text = str_content;
                                }
                            }
                        }
                    }
                    remove_replace_with_white_space(out replaceKey, out replaceValue, table);
                    Paragraph paragraph_nextpage = doc.Paragraphs.Add();
                    paragraph_nextpage.Range.InsertBreak();

                    if ((last_percent) < ((row - 1) * 100) / cell_max_row)
                    {
                        last_percent = ((row - 1) * 100) / cell_max_row;
                        backgroundWorker.ReportProgress(last_percent, "处理第"+row+"行");
                    }
                }
                object fileformat = WdSaveFormat.wdFormatDocument;
                doc.SaveAs2(outputfullfilename);
            }
            finally
            {
                object doNotSaveChanges = WdSaveOptions.wdDoNotSaveChanges;
                doc?.Close(ref doNotSaveChanges, ref miss, ref miss);
                doc_template?.Close(ref doNotSaveChanges, ref miss, ref miss);
                appWord.Application.Quit(ref miss, ref miss, ref miss);
                doc = null;
                doc_template = null;
                appWord = null;
            }
        }
    }
}
