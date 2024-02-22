using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;

namespace LecturaExcel.Responsabilitis
{
    public class Diferential
    {
        public void test(DataGridView dgvToReview, DataGridView dgvToAddValue)
        {
            string name;
            double cellValue;
            double referenceValue = 100000;
            foreach (DataGridViewRow row in dgvToReview.Rows)
            {
                cellValue = Convert.ToDouble(row.Cells[2].Value.ToString());
                if (cellValue < referenceValue)
                {
                    referenceValue = cellValue;
                    name = row.Cells[0].Value.ToString();
                    dgvToAddValue.Rows[0].Cells[0].Value = name;
                }
            }
            dgvToAddValue.Rows[0].Cells[1].Value = Math.Round(referenceValue, 2);
        }

        public void assignComparisonVariable(DataGridView dgvToReview, DataGridView dgvToAddVarible)
        {
            double cellValue;
            double referenceValue = 100000;
            string name;
            try
            {
                foreach (DataGridViewRow row in dgvToReview.Rows)
                {
                    cellValue = Convert.ToDouble(row.Cells[2].Value.ToString());
                    if (cellValue < referenceValue)
                    {
                        referenceValue = cellValue;
                        name = row.Cells[0].Value.ToString();
                        dgvToAddVarible.Rows[0].Cells[0].Value = name;
                    }
                }
                dgvToAddVarible.Rows[0].Cells[1].Value = Math.Round(referenceValue, 2);
            }
            catch (Exception ex)
            {
                Trace.WriteLine(ex.ToString());
            }
        }


    }
}
