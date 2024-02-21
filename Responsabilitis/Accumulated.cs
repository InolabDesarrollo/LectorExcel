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
        private readonly DataGridView dataGridViewWithData;

        public Accumulated()
        {

        }
        public Accumulated(DataGridView dataGridViewWithData)
        {
            this.dataGridViewWithData = dataGridViewWithData;
        }

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

        public void addAccumulatedValuesToRightRunOne(DataGridView dgvToReview, DataGridView dgvToAddAccumulated)
        {
            foreach (DataGridViewRow row in dgvToReview.Rows)
            {
                double accumulated = 0;
                int cellValue = Convert.ToInt32(row.Cells[3].Value);
                int n = 1;
                //aumentar a la fila los valores acumulativos a la derecha (los que van arriba)
                try
                {
                    while (n <= cellValue)
                    {
                        accumulated = accumulated + Convert.ToDouble(dataGridViewWithData.Rows[n].Cells[2].Value);
                        n++;
                        if (accumulated > 100)
                        {
                            accumulated = 100;
                        }
                        dgvToAddAccumulated.Rows[row.Index].Cells[2].Value = Math.Round(accumulated, 2);
                    }
                }
                catch (Exception ex)
                {
                    Trace.WriteLine(ex.Message);
                }
            }           
        }

        public void addAccumulatedValuesToRightRunTwo(DataGridView dgvToReview, DataGridView dgvToAddAccumulated)
        {
            foreach (DataGridViewRow row in dgvToReview.Rows)
            {
                double accumulated = 0;
                int cellValue = Convert.ToInt32(row.Cells[3].Value);
                int n = 1;
                //aumentar a la fila los valores acumulativos a la derecha (los que van arriba)
                try
                {
                    while (n <= cellValue)
                    {
                        accumulated = accumulated + Convert.ToDouble(dataGridViewWithData.Rows[n].Cells[3].Value);
                        n++;
                        if (accumulated > 100)
                        {
                            accumulated = 100;
                        }
                        dgvToAddAccumulated.Rows[row.Index].Cells[3].Value = Math.Round(accumulated, 2);
                    }
                }
                catch (Exception ex)
                {
                    Trace.WriteLine(ex.Message);
                }
            }
        }

        public void addAccumulatedValuesToRightRunThree(DataGridView dgvToReview, DataGridView dgvToAddAccumulated, int numberOfCellToAddAccumulated)
        {
            foreach (DataGridViewRow row in dgvToReview.Rows)
            {
                double accumulated = 0;
                int cellValue = Convert.ToInt32(row.Cells[3].Value);
                int num = Convert.ToInt32(row.Cells[3].Value) + 1;
                try
                {
                    while (num > cellValue)
                    {
                        accumulated = accumulated + Convert.ToDouble(dataGridViewWithData.Rows[num].Cells[2].Value);
                        num++;
                        if (accumulated > 100)
                        {
                            accumulated = 100;
                        }
                        dgvToAddAccumulated.Rows[row.Index].Cells[numberOfCellToAddAccumulated].Value = Math.Round(accumulated, 2);
                    }
                }
                catch (Exception ex)
                {
                    Trace.WriteLine(ex.Message);
                }
            }
        }

        public void addAccumulatedValuesToRightRunFor(DataGridView dgvToReview, DataGridView dgvToAddAccumulated)
        {
            foreach (DataGridViewRow row in dgvToReview.Rows)
            {
                double accumulated = 0;
                int cellValue = Convert.ToInt32(row.Cells[3].Value);
                int num = Convert.ToInt32(row.Cells[3].Value) + 1;
                try
                {
                    while (num > cellValue)
                    {
                        accumulated = accumulated + Convert.ToDouble(dataGridViewWithData.Rows[num].Cells[3].Value);
                        num++;
                        if (accumulated > 100)
                        {
                            accumulated = 100;
                        }
                        dgvToAddAccumulated.Rows[row.Index].Cells[5].Value = Math.Round(accumulated, 2);
                    }
                }
                catch (Exception ex)
                {
                    Trace.WriteLine(ex.Message);
                }
            }
        }

        public void addCumulativeValuesToLeftBy100(DataGridView dgvToReview, DataGridView dgvWithAccumulativeValues)
        {
            foreach (DataGridViewRow row in dgvToReview.Rows)
            //Llenado de los acumulativos a la izquierda por medio de total a 100 
            {
                double resultado = 100 - Convert.ToDouble(dgvWithAccumulativeValues.Rows[row.Index].Cells[2].Value);
                dgvWithAccumulativeValues.Rows[row.Index].Cells[4].Value = Math.Round(resultado, 2);

                double resultado2 = 100 - Convert.ToDouble(dgvWithAccumulativeValues.Rows[row.Index].Cells[3].Value);
                dgvWithAccumulativeValues.Rows[row.Index].Cells[5].Value = Math.Round(resultado2, 2);
            }
        }

        public void addCumulativeValuesToLeftBy100RunOne(DataGridView dgvToReview, DataGridView dgvWithAccumulativeValues)
        {
            foreach (DataGridViewRow row in dgvToReview.Rows)
            {
                double resultado = 100 - Convert.ToDouble(dgvWithAccumulativeValues.Rows[row.Index].Cells[2].Value);
                dgvWithAccumulativeValues.Rows[row.Index].Cells[3].Value = Math.Round(resultado, 2);
            }
        }


    }
}