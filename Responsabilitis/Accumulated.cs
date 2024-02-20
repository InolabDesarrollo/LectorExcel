using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static OfficeOpenXml.ExcelErrorValue;

namespace LecturaExcel.Responsabilitis
{
    public class Accumulated
    {
      
        public void addCumulativeValuesToRight100(DataGridView dgvToFill, DataGridView dgvWithAccumulatedValues)
        {
            foreach (DataGridViewRow row in dgvToFill.Rows)
            {
                double resultado = 100 - Convert.ToDouble(dgvWithAccumulatedValues.Rows[row.Index].Cells[2].Value);
                dgvWithAccumulatedValues.Rows[row.Index].Cells[5].Value = Math.Round(resultado, 2);
                dgvToFill.Rows[row.Index].Cells[5].Value = Math.Round(resultado, 2);

                double resultado2 = 100 - Convert.ToDouble(dgvWithAccumulatedValues.Rows[row.Index].Cells[3].Value);
                dgvWithAccumulatedValues.Rows[row.Index].Cells[6].Value = Math.Round(resultado2, 2);
                dgvToFill.Rows[row.Index].Cells[6].Value = Math.Round(resultado2, 2);

                double resultado3 = 100 - Convert.ToDouble(dgvWithAccumulatedValues.Rows[row.Index].Cells[4].Value);
                dgvWithAccumulatedValues.Rows[row.Index].Cells[7].Value = Math.Round(resultado3, 2);
                dgvToFill.Rows[row.Index].Cells[7].Value = Math.Round(resultado3, 2);
            }

        }
    }
}