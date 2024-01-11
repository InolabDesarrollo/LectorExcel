using Microsoft.Reporting.WinForms;
using System;
using System.Data;
using System.Windows.Forms;

namespace LecturaExcel
{
    public partial class Vista_i : Form
    {        
        public Vista_i(DataTable dt)
        {
            InitializeComponent();
            dt1 = dt;
        }
        DataTable dt1 = new DataTable();

        private void Vista_i_Load(object sender, EventArgs e)
        {
            //De donde sacara la informacion para llenar el reporte
            WindowState = FormWindowState.Maximized;
            reportViewer1.LocalReport.DataSources.Clear();
            reportViewer1.LocalReport.DataSources.Add(new ReportDataSource("mydata", dt1));
            this.reportViewer1.RefreshReport();

        }
    }
}

