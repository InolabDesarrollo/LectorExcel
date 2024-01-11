using Microsoft.Reporting.WinForms;
using System;
using System.Data;
using System.Windows.Forms;

namespace LecturaExcel
{
    public partial class Vista2 : Form
    {
        public Vista2(DataTable dt)
        {
            InitializeComponent();
            dt1 = dt;
        }
        DataTable dt1 = new DataTable();

        private void Vista2_Load(object sender, EventArgs e)
        {
            //De donde sacara los datos para el reporte
            WindowState = FormWindowState.Maximized;
            reportViewer1.LocalReport.DataSources.Clear();
            reportViewer1.LocalReport.DataSources.Add(new ReportDataSource("mydata", dt1));
            this.reportViewer1.RefreshReport();
        }
    }
}
