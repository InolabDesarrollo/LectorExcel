using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LecturaExcel.Responsabilitis
{
    public class Differential
    {
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

        public void assignComparisonVariable(DataGridView dgvToReview, DataGridView dgvToAddVarible, int cellWithValue, int cellToAddValue)
        {
            double cellValue;
            double referenceValue = 100000;
            string name;
            try
            {
                foreach (DataGridViewRow row in dgvToReview.Rows)
                {
                    cellValue = Convert.ToDouble(row.Cells[cellWithValue].Value.ToString());
                    if (cellValue < referenceValue)
                    {
                        referenceValue = cellValue;
                        name = row.Cells[0].Value.ToString();
                        dgvToAddVarible.Rows[0].Cells[0].Value = name;
                    }
                }
                dgvToAddVarible.Rows[0].Cells[cellToAddValue].Value = Math.Round(referenceValue, 2);
            }
            catch (Exception ex)
            {
                Trace.WriteLine(ex.ToString());
            }
        }

        public void assignComparisonVariableRowTwo(DataGridView dgvToReview, DataGridView dgvToAddVarible)
        {
            double cellValue;
            double referenceValue = 1000000;
            string name;
            try
            {
                foreach (DataGridViewRow row in dgvToReview.Rows)
                {
                    cellValue = Convert.ToDouble(row.Cells[3].Value.ToString());
                    if (cellValue < referenceValue)
                    {
                        referenceValue = cellValue;
                        name = row.Cells[0].Value.ToString();
                        dgvToAddVarible.Rows[2].Cells[0].Value = name;
                    }
                }
                dgvToAddVarible.Rows[2].Cells[1].Value = Math.Round(referenceValue, 2);
            }
            catch (Exception ex)
            {
                Trace.WriteLine(ex.ToString());
            }
        }

        public void assignComparisonVariableRowTwo(DataGridView dgvToReview, DataGridView dgvToAddVarible, int cellWithValue, int cellToAddValue)
        {
            double cellValue;
            double referenceValue = 1000000;
            string name;
            try
            {
                foreach (DataGridViewRow row in dgvToReview.Rows)
                {
                    cellValue = Convert.ToDouble(row.Cells[cellWithValue].Value.ToString());
                    if (cellValue < referenceValue)
                    {
                        referenceValue = cellValue;
                        name = row.Cells[0].Value.ToString();
                        dgvToAddVarible.Rows[2].Cells[0].Value = name;
                    }
                }
                dgvToAddVarible.Rows[2].Cells[cellToAddValue].Value = Math.Round(referenceValue, 2);
            }
            catch (Exception ex)
            {
                Trace.WriteLine(ex.ToString());
            }
        }

        public void createDifferential(DataGridView dgvWithDifferentialValues)
        {
            dgvWithDifferentialValues.Rows[1].Cells[1].Value =
            Math.Round(Convert.ToDouble(100 - (Convert.ToDouble(dgvWithDifferentialValues.Rows[2].Cells[1].Value) +
            Convert.ToDouble(dgvWithDifferentialValues.Rows[0].Cells[1].Value))), 2);
        }

        public void assignComparisonVariableRunTwo(DataGridView dgvToReview, DataGridView dgvToAddVarible)
        {
            double cellValue;
            double referenceValue = 100000;
            foreach (DataGridViewRow row in dgvToReview.Rows)
            {
                cellValue = Convert.ToDouble(row.Cells[3].Value.ToString());
                if (cellValue < referenceValue)
                {
                    referenceValue = cellValue;
                }
            }
            dgvToAddVarible.Rows[0].Cells[2].Value = Math.Round(referenceValue, 2);
        }

        public void assignComparisonVariableRunTwo(DataGridView dgvToReview, DataGridView dgvToAddVarible, int cellWithValue, int cellToAddValue)
        {
            double cellValue;
            double referenceValue = 100000;
            foreach (DataGridViewRow row in dgvToReview.Rows)
            {
                cellValue = Convert.ToDouble(row.Cells[cellWithValue].Value.ToString());
                if (cellValue < referenceValue)
                {
                    referenceValue = cellValue;
                }
            }
            dgvToAddVarible.Rows[0].Cells[cellToAddValue].Value = Math.Round(referenceValue, 2);
        }

    }
}
