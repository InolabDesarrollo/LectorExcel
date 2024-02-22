using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LecturaExcel.Responsabilitis
{
    public class Manage_Data
    {

        public void createDgvToleranceTableReference(DataGridView DgvToFill)
        {
            DgvToFill.Rows.Add("20", "635", "29µm", "35µm");
            DgvToFill.Rows.Add("25", "500", "34µm", "41µm");
            DgvToFill.Rows.Add("32", "450", "42µm", "50µm");
            DgvToFill.Rows.Add("38", "400", "48µm", "57µm");
            DgvToFill.Rows.Add("45", "325", "57µm", "66µm");
            DgvToFill.Rows.Add("53", "270", "66µm", "76µm");
            DgvToFill.Rows.Add("63", "230", "77µm", "89µm");
            DgvToFill.Rows.Add("75", "200", "91µm", "103µm");
            DgvToFill.Rows.Add("90", "170", "108µm", "122µm");
            DgvToFill.Rows.Add("106", "140", "126µm", "141µm");
            DgvToFill.Rows.Add("125", "120", "147µm", "163µm");
            DgvToFill.Rows.Add("150", "100", "174µm", "192µm");
            DgvToFill.Rows.Add("180", "80", "207µm", "227µm");
            DgvToFill.Rows.Add("212", "70", "242µm", "263µm");
            DgvToFill.Rows.Add("250", "60", "283µm", "306µm");
            DgvToFill.Rows.Add("300", "50", "337µm", "363µm");
            DgvToFill.Rows.Add("355", "45", "396µm", "425µm");
            DgvToFill.Rows.Add("425", "40", "471µm", "502µm");
            DgvToFill.Rows.Add("500", "35", "550µm", "585µm");
            DgvToFill.Rows.Add("600", "30", "660µm", "695µm");
            DgvToFill.Rows.Add("710", "25", "775µm", "815µm");
            DgvToFill.Rows.Add("850", "20", "925µm", "970µm");
            DgvToFill.Rows.Add("1000", "18", "1.083mm", "1.135mm");
            DgvToFill.Rows.Add("1180", "16", "1.270mm", "1.330mm");
            DgvToFill.Rows.Add("1400", "14", "1.505mm", "1.565mm");
            DgvToFill.Rows.Add("1700", "12", "1.820mm", "1.890mm");
            DgvToFill.Rows.Add("2000", "10", "2.135mm", "2.215mm");
            DgvToFill.Rows.Add("2360", "8", "2.515mm", "2.609mm");
            DgvToFill.Rows.Add("2800", "7", "2.975mm", "3.070mm");
            DgvToFill.Rows.Add("3350", "6", "3.55mm", "3.66mm");
        }

        public void addColumnToDatagridView(string headerText, DataGridView dataGridView)
        {
            DataGridViewTextBoxColumn column = new DataGridViewTextBoxColumn();
            column.HeaderText = headerText;
            column.Width = 100;
            dataGridView.Columns.Add(column);
        }

        public void copyStructureOfDataGridViewToOther(DataGridView Dgv_Original, DataGridView Dgv_ToCopy)
        {
            try
            {
                foreach (DataGridViewColumn column in Dgv_Original.Columns)
                {
                    if (column.HeaderText != "")
                    {
                        DataGridViewTextBoxColumn columnNew = new DataGridViewTextBoxColumn();
                        columnNew.HeaderText = column.Name;
                        columnNew.Width = 100;
                        Dgv_ToCopy.Columns.Add(columnNew);
                    }
                }

                Dgv_ToCopy.Rows.Add(Dgv_Original.Rows[0].Cells[0].Value.ToString(),
                    Dgv_Original.Rows[0].Cells[2].Value.ToString(), Dgv_Original.Rows[0].Cells[3].Value.ToString(),
                    Dgv_Original.Rows[0].Cells[4].Value.ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error  copyStructureOfDataGridViewToOther" + ex.Message);
            }
        }

        public void removeUselessGridColumns(DataGridView dataGridView)
        {
            try
            {
                if (dataGridView.Columns.Count >= 3)
                {
                    for (int i = 2; i <= dataGridView.Columns.Count; i++)
                    {
                        dataGridView.Columns.RemoveAt(i);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error removeUselessGridColumns" + ex.Message);
            }
        }

        public void hideCummulativeValues(string numberOfRuns, DataGridView dgvToHide)
        {
            try
            {
                if (numberOfRuns == "3")
                {
                    this.hideColumnsOfDataGridView(dgvToHide, 5);
                }
                else if (numberOfRuns == "2")
                {
                    this.hideColumnsOfDataGridView(dgvToHide, 4);
                }
                else if (numberOfRuns == "1")
                {
                    this.hideColumnsOfDataGridView(dgvToHide, 3);
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show("Columns have been removed " + ex.Message);
            }       
        }

        private void hideColumnsOfDataGridView(DataGridView dgvToHide, int columnIndex )
        {
            foreach (DataGridViewRow row in dgvToHide.Rows)
            {
                foreach (DataGridViewColumn col in dgvToHide.Columns)
                {
                    if (col.Index >= columnIndex)
                    {
                        dgvToHide.Rows[row.Index].Cells[col.Index].Value = "";
                    }
                }
            }
        }

        public void addRowsToDataGridView(DataGridView dgvToAddRows)
        {
            for (int i=0; i<3; i++)
            {
                dgvToAddRows.Rows.Add();
            }
        }
    }
}
