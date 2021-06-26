using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelChartCtrl
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();

            this.TopMost = true;

            clsCommon.SetDir();
        }

        private void BTN_Debug_Click(object sender, EventArgs e)
        {
            var xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbooks xlBooks;

            xlBooks = xlApp.Workbooks;
            xlBooks.Open(clsCommon.strExcelFileName);

            xlApp.Visible = true;

            string strT = "";
            xlApp.Run("ChartCtrl_0_8_1.xlsm!test", strT);

            MessageBox.Show(strT);
        }
    }
}
