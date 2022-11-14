using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TestCaseAndProblem
{
    public class RangeUtils
    {
        public static int get_max_column(Range Cells)
        {
            Range rng = Cells[1, Cells.Columns.Count];
            return rng.get_End(XlDirection.xlToLeft).Column;
        }

        public static int get_max_row(Range Cells)
        {
            Range rng = Cells[Cells.Rows.Count, 1];
            return rng.get_End(XlDirection.xlUp).Row;
        }

        public static int get_row_by_title(Range Cells, String title, int column = 1) {
            int max_row = Cells.Rows.Count;
            for (int i = 1; i <= max_row; i++) {
                if (String.Equals(Cells[i, column].Text, title)) {
                    return i;
                }
            }
            return 0;
        }

        public static int get_column_by_title(Range Cells, String title, int row = 1)
        {
            int max_column = get_max_column(Cells);
            for (int i = 1; i <= max_column; i++)
            {
                if (String.Equals(Cells[row, i].Text, title))
                {
                    return i;
                }
            }
            return 0;
        }

        public static int get_column_by_title(Range Cells, int max_column, String title, int row = 1)
        {
            for (int i = 1; i <= max_column; i++)
            {
                if (String.Equals(Cells[row, i].Text, title))
                {
                    return i;
                }
            }
            return 0;
        }
    }
}
