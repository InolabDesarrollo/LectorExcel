using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LecturaExcel.Responsabilitis
{
    public class Accumulated
    {
        public void addAccumulatedToRightBy100(DataGridView DgvToFill, DataGridView DgvWithAccumulatedValues)
        {
            try
            {
                foreach (DataGridViewRow row in DgvToFill.Rows)
                {
                    int cellWithValue = 2;
                    for (int i = 5; i == 7; i++)
                    {
                        double accumulated = 100 - Convert.ToDouble(DgvWithAccumulatedValues.Rows[row.Index].Cells[cellWithValue].Value);

                        DgvWithAccumulatedValues.Rows[row.Index].Cells[i].Value = Math.Round(accumulated, 2);
                        DgvToFill.Rows[row.Index].Cells[i].Value = Math.Round(accumulated, 2);
                        cellWithValue++;
                    }

                }
            }
            catch (Exception ex)
            {
                Trace.WriteLine(ex.ToString());
            }
        }
      
    }
}