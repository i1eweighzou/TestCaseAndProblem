using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;

namespace TestCaseAndProblem
{
    partial class ActionsPaneControlTestCase : UserControl
    {
        internal Range range;
        private InterfaceUpdateText interfaceUpdateText;
        public ActionsPaneControlTestCase()
        {
            InitializeComponent();
        }

        public InterfaceUpdateText InterfaceUpdateText { get => interfaceUpdateText; set => interfaceUpdateText = value; }

        public void set_range(Range sel_range) { 
            if (sel_range != null && sel_range.Cells.Count == 1)
            {
                range = sel_range;
                textBoxDetail.Text = range.Text;
            }
            else {
                range = null;
            }            
        }

        public void show_progress(int step, String info) {
            System.Action AsyncUIDelegate = delegate () {
            };
            this.Invoke(AsyncUIDelegate);           
        }

        private void textBoxDetail_Enter(object sender, EventArgs e)
        {
            interfaceUpdateText?.selelect_lost_focus();
        }

        private void textBoxDetail_Leave(object sender, EventArgs e)
        {
            if (range != null) {
                range.Cells[1, 1].Value = textBoxDetail.Text;
            }
            
            interfaceUpdateText?.selelect_focus(textBoxDetail.Text);
        }       
    }
}
