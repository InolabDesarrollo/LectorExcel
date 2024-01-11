using System;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using System.IO;
using ExcelDataReader;
using System.Text.RegularExpressions;
using System.Collections.Generic;

namespace LecturaExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

        }
        //Declaracion de variables 
        string filaecu;
        string filaecu2;
        string Valor;
        string Valor2;
        String strFila;
        String strFila2;
        string name;
        string namez;
        string lname;
        string lnamez;
        string tabla = "";
        int check = 0;
        int corridas = 3;

        bool ch1 = true;
        bool ch2 = true;
        bool allowSelect = false;
        bool oc3 = true;
        bool oc1 = true;
        bool oc2 = true;
        string Ace1 = "";
        string Ace2 = "";
        string Ace3 = "";

        string num_corr = "";
        string con_ocu = "no";
        string con_dif = "no";

        //Listas de valores de datos de empresas 
        List<string> Nombres = new List<string>();
        List<string> Fecha = new List<string>();
        List<string> Usuarios = new List<string>();
        List<string> Equipos = new List<string>();
        List<string> Ids = new List<string>();
        List<string> Grupos = new List<string>();
        List<string> Lotes = new List<string>();
        List<string> Comentarios = new List<string>();
        List<string> Clientes = new List<string>();
        List<string> Ase = new List<string>();

        List<string> valor_nominal= new List<string>();

        private void Form1_Load(object sender, EventArgs e)
        {
            //Maximizar el tamaño de la ventana del form
            WindowState = FormWindowState.Maximized;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Se hace la subida de un archivo de excel ccon las especificaciones de laboratorios Pisa
            OpenFileDialog fil = new OpenFileDialog();
            fil.ShowDialog();
            string path = fil.FileName.ToString();
            try
            {
                ExcelFileReader(path);
            }
            catch (Exception es)
            {
                MessageBox.Show("Please select an excel file with the correct format");
            }          
            dataGridView1.ReadOnly = true;
            datos.ReadOnly = true;
        }

        public void ExcelFileReader(string path)
        {
            //Se hace la lectura del Excel para las hojas de nombre "Data" y "Sample Info" para sus Grid correspondientes
            var stream = File.Open(path, FileMode.Open, FileAccess.Read);
            var reader = ExcelReaderFactory.CreateReader(stream);
            var result = reader.AsDataSet();
            var tables = result.Tables.Cast<DataTable>();
            foreach (DataTable table in tables)
            {
                if (table.ToString() == "Data")
                {
                    dataGridView1.DataSource = table;
                    datos.DataSource = table;
                }
                if (table.ToString() == "Sample Info")
                {
                    dataGridView7.DataSource = table;
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            valor_nominal.Add("-");
            try
            {
                allowSelect = true;
                tabControl1.SelectedTab = tabPage2;
                allowSelect = false;

                //Tabla de valores de referencia para ejecutar el analisis de calculos
                Referencias.Rows.Add("20", "635", "29µm", "35µm");
                Referencias.Rows.Add("25", "500", "34µm", "41µm");
                Referencias.Rows.Add("32", "450", "42µm", "50µm");
                Referencias.Rows.Add("38", "400", "48µm", "57µm");
                Referencias.Rows.Add("45", "325", "57µm", "66µm");
                Referencias.Rows.Add("53", "270", "66µm", "76µm");
                Referencias.Rows.Add("63", "230", "77µm", "89µm");
                Referencias.Rows.Add("75", "200", "91µm", "103µm");
                Referencias.Rows.Add("90", "170", "108µm", "122µm");
                Referencias.Rows.Add("106", "140", "126µm", "141µm");
                Referencias.Rows.Add("125", "120", "147µm", "163µm");
                Referencias.Rows.Add("150", "100", "174µm", "192µm");
                Referencias.Rows.Add("180", "80", "207µm", "227µm");
                Referencias.Rows.Add("212", "70", "242µm", "263µm");
                Referencias.Rows.Add("250", "60", "283µm", "306µm");
                Referencias.Rows.Add("300", "50", "337µm", "363µm");
                Referencias.Rows.Add("355", "45", "396µm", "425µm");
                Referencias.Rows.Add("425", "40", "471µm", "502µm");
                Referencias.Rows.Add("500", "35", "550µm", "585µm");
                Referencias.Rows.Add("600", "30", "660µm", "695µm");
                Referencias.Rows.Add("710", "25", "775µm", "815µm");
                Referencias.Rows.Add("850", "20", "925µm", "970µm");
                Referencias.Rows.Add("1000", "18", "1.083mm", "1.135mm");
                Referencias.Rows.Add("1180", "16", "1.270mm", "1.330mm");
                Referencias.Rows.Add("1400", "14", "1.505mm", "1.565mm");
                Referencias.Rows.Add("1700", "12", "1.820mm", "1.890mm");
                Referencias.Rows.Add("2000", "10", "2.135mm", "2.215mm");
                Referencias.Rows.Add("2360", "8", "2.515mm", "2.609mm");
                Referencias.Rows.Add("2800", "7", "2.975mm", "3.070mm");
                Referencias.Rows.Add("3350", "6", "3.55mm", "3.66mm");

                //Añadir los campos de "Nominal Sieve size" y "Mesh#"
                DataGridViewTextBoxColumn columna1 = new DataGridViewTextBoxColumn();
                columna1.HeaderText = "Nominal Sieve Size";
                columna1.Width = 100;

                DataGridViewTextBoxColumn columna2 = new DataGridViewTextBoxColumn();
                columna2.HeaderText = "Mesh #";
                columna2.Width = 70;

                DataGridViewTextBoxColumn columna3 = new DataGridViewTextBoxColumn();
                columna3.HeaderText = "Values To Calculate";
                columna3.Width = 70;

                DataGridViewTextBoxColumn columna4 = new DataGridViewTextBoxColumn();
                columna3.HeaderText = "Record";
                columna3.Width = 50;

                dataGridView2.Columns.Add(columna1);
                dataGridView2.Columns.Add(columna2);
                dataGridView2.Columns.Add(columna3);
                dataGridView2.Columns.Add(columna4);

                //Añadir los campos de "Nominal Sieve size" y "Mesh#"
                DataGridViewTextBoxColumn columna11 = new DataGridViewTextBoxColumn();
                columna11.HeaderText = "Nominal Sieve Size";
                columna11.Width = 100;

                DataGridViewTextBoxColumn columna21 = new DataGridViewTextBoxColumn();
                columna21.HeaderText = "Mesh #";
                columna21.Width = 70;

                DataGridViewTextBoxColumn columna31 = new DataGridViewTextBoxColumn();
                columna31.HeaderText = "Values To Calculate";
                columna31.Width = 70;

                DataGridViewTextBoxColumn columna41 = new DataGridViewTextBoxColumn();
                columna31.HeaderText = "Record";
                columna31.Width = 50;

                dataGridView13.Columns.Add(columna11);
                dataGridView13.Columns.Add(columna21);
                dataGridView13.Columns.Add(columna31);
                dataGridView13.Columns.Add(columna41);

                //llenar FilaSeleccoinada con las columnas que son
                foreach (DataGridViewColumn cols in dataGridView1.Columns)
                {
                    if (cols.HeaderText != "")
                    {
                        //Añadir los campos de "Columna" del datagrid1
                        DataGridViewTextBoxColumn columnas = new DataGridViewTextBoxColumn();
                        columnas.HeaderText = cols.Name;
                        columnas.Width = 100;
                        FilaSeleccionada.Columns.Add(columnas);
                    }
                }

                //Agregar la primera fila del datagrid1 a Fila Seleccionada
                FilaSeleccionada.Rows.Add(dataGridView1.Rows[0].Cells[0].Value.ToString(), 
                    dataGridView1.Rows[0].Cells[2].Value.ToString(), dataGridView1.Rows[0].Cells[3].Value.ToString(), 
                    dataGridView1.Rows[0].Cells[4].Value.ToString());

                //llenar FilaSeleccoinada2 con las columnas que son
                foreach (DataGridViewColumn cols in dataGridView1.Columns)
                {
                    if (cols.HeaderText != "")
                    {
                        //Añadir los campos de "Columna" del datagrid1
                        DataGridViewTextBoxColumn columnas1 = new DataGridViewTextBoxColumn();
                        columnas1.HeaderText = cols.Name;
                        columnas1.Width = 100;
                        FilaSeleccionada2.Columns.Add(columnas1);
                    }
                }

                //Agregar la primera fila del datagrid1 a Fila Seleccionada
                FilaSeleccionada2.Rows.Add(dataGridView1.Rows[0].Cells[0].Value.ToString(),
                    dataGridView1.Rows[0].Cells[2].Value.ToString(), dataGridView1.Rows[0].Cells[3].Value.ToString(),
                    dataGridView1.Rows[0].Cells[4].Value.ToString());
            }
            catch (Exception f)
            {
                MessageBox.Show("Please select an Excel file to continue");
            }
            //Limpieza de loa que habia antes de los GridCorrespondientes
            dataGridView2.ReadOnly = true;
            dataGridView13.ReadOnly = true;
            dataGridView2.Rows.Clear();
            dataGridView3.Rows.Clear();
            dataGridView4.Rows.Clear();
            dataGridView5.Rows.Clear();
            dataGridView11.Rows.Clear();
            dataGridView6.Rows.Clear();
            dataGridView12.Rows.Clear();
            dataGridView14.Rows.Clear();
            dataGridView15.Rows.Clear();

            ch1 = true;
            ch2 = true;

            try
            {
                while (true)
                {
                    //remueve las columnas de el grid 5 desde l 3er columna en adelante
                    dataGridView5.Columns.RemoveAt(2);
                }
            }
            catch (Exception l)
            {

            }
            try
            {
                while (true)
                {
                    //remueve las columnas de el grid 5 desde l 3er columna en adelante
                    dataGridView3.Columns.RemoveAt(2);
                }
            }
            catch (Exception l)
            {

            }
            try
            {
                while (true)
                {
                    //remueve las columnas de el grid 5 desde l 3er columna en adelante
                    dataGridView11.Columns.RemoveAt(2);
                }
            }
            catch (Exception l)
            {

            }
            try
            {
                while (true)
                {
                    //remueve las columnas de el grid 5 desde l 3er columna en adelante
                    dataGridView14.Columns.RemoveAt(2);
                }
            }
            catch (Exception l)
            {

            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Despues de seleccionar una malla que se quiera conocer sus datos dentro del excel, se hace una lista de referencia correspondiendo a lo que hay en el grid de el excel que se subio y con los datos de referencia que marcan los limites de cada malla
            dataGridView2.Rows.Add(Referencias.Rows[comboBox1.SelectedIndex].Cells[2].Value.ToString(), 
                comboBox1.SelectedItem.ToString(), Referencias.Rows[comboBox1.SelectedIndex].Cells[2].Value.ToString(), true);
            dataGridView13.Rows.Add(Referencias.Rows[comboBox1.SelectedIndex].Cells[3].Value.ToString(),
                comboBox1.SelectedItem.ToString(), Referencias.Rows[comboBox1.SelectedIndex].Cells[3].Value.ToString(), true);
            dataGridView3.Rows.Add(Referencias.Rows[comboBox1.SelectedIndex].Cells[2].Value.ToString(), 
                comboBox1.SelectedItem.ToString(), Referencias.Rows[comboBox1.SelectedIndex].Cells[2].Value.ToString(), true);
            dataGridView14.Rows.Add(Referencias.Rows[comboBox1.SelectedIndex].Cells[3].Value.ToString(),
                comboBox1.SelectedItem.ToString(), Referencias.Rows[comboBox1.SelectedIndex].Cells[3].Value.ToString(), true);
            dataGridView5.Rows.Add(Referencias.Rows[comboBox1.SelectedIndex].Cells[2].Value.ToString(), 
                comboBox1.SelectedItem.ToString(), Referencias.Rows[comboBox1.SelectedIndex].Cells[2].Value.ToString(), true);
            dataGridView11.Rows.Add(Referencias.Rows[comboBox1.SelectedIndex].Cells[3].Value.ToString(),
                comboBox1.SelectedItem.ToString(), Referencias.Rows[comboBox1.SelectedIndex].Cells[3].Value.ToString(), true);

            //Configuracion de los Grid
            dataGridView2.AllowUserToAddRows = false;
            dataGridView13.AllowUserToAddRows = false;
            dataGridView3.AllowUserToAddRows = false;
            dataGridView14.AllowUserToAddRows = false;

            //Se limpia el campo del listbox
            FilaSeleccionada.Rows.Clear();
            foreach (DataGridViewColumn cols in dataGridView1.Columns)
            {
                if (cols.HeaderText != "")
                {
                    //Añadir los campos de "Columna" del datagrid1
                    DataGridViewTextBoxColumn columnas = new DataGridViewTextBoxColumn();
                    columnas.HeaderText = cols.Name;
                    columnas.Width = 100;
                    FilaSeleccionada.Columns.Add(columnas);
                }
                else
                {

                }
            }
            FilaSeleccionada2.Rows.Clear();
            foreach (DataGridViewColumn cols in dataGridView1.Columns)
            {
                if (cols.HeaderText != "")
                {
                    //Añadir los campos de "Columna" del datagrid1
                    DataGridViewTextBoxColumn columnas1 = new DataGridViewTextBoxColumn();
                    columnas1.HeaderText = cols.Name;
                    columnas1.Width = 100;
                    FilaSeleccionada2.Columns.Add(columnas1);
                }
                else
                {

                }
            }
            //Agregar la primera fila del datagrid1 a Fila Seleccionada
            FilaSeleccionada.Rows.Add(dataGridView1.Rows[0].Cells[0].Value.ToString(), dataGridView1.Rows[0].Cells[2].Value.ToString(), dataGridView1.Rows[0].Cells[3].Value.ToString(), dataGridView1.Rows[0].Cells[4].Value.ToString());
            FilaSeleccionada2.Rows.Add(dataGridView1.Rows[0].Cells[0].Value.ToString(), dataGridView1.Rows[0].Cells[2].Value.ToString(), dataGridView1.Rows[0].Cells[3].Value.ToString(), dataGridView1.Rows[0].Cells[4].Value.ToString());
            //Para dgv2
            try
            {
                //Ya que hace la busqueda del valor mas cercano al de la lista de referencia lo coloca en el Grid2
                foreach (DataGridViewRow Row in dataGridView2.Rows)
                {
                    check = 0;

                    string texto = Convert.ToString(Row.Cells[2].Value);
                    Match m = Regex.Match(texto, "(\\d+)");
                    string num = string.Empty;
                    if (m.Success)
                    {
                        num = m.Value;
                    }
                    double h = Convert.ToDouble(num);
                    double f = Math.Round(h, 0);
                    if ((f == 1) || (f == 2) || (f == 3) || (f == 4))
                    {
                        //Bugfix de cuando tiene enteros
                        double r = h;
                        r = r * 1000;
                        r = Math.Round(r, 0);
                        BusquedaT(r.ToString());
                        //Busqueda hasta que encuentre un valor aproximado al que hay en la lista de referencia
                        double val2 = r;
                        while (check == 0)
                        {

                            val2 = val2 - 1;
                            //Busqueda al valor aproximado menor
                            BusquedaT(val2.ToString());
                        }
                    }
                    else
                    {
                        Busqueda(f.ToString());
                        //Busqueda hasta que encuentre un valor aproximado al que hay en la lista de referencia
                        double val2 = f;
                        while (check == 0)
                        {

                            val2 = val2 - 1;
                            //Busqueda al valor aproximado menor
                            Busqueda(val2.ToString());
                        }
                    }
                    Row.Cells[3].Value = filaecu;
                }
            }
            catch (Exception f)
            {

            }
            //Para dgv13
            try
            {
                //Ya que hace la busqueda del valor mas cercano al de la lista de referencia lo coloca en el Grid13
                foreach (DataGridViewRow Row in dataGridView13.Rows)
                {
                    check = 0;

                    string texto = Convert.ToString(Row.Cells[2].Value);
                    Match m = Regex.Match(texto, "(\\d+)");
                    string num = string.Empty;
                    if (m.Success)
                    {
                        num = m.Value;
                    }
                    double h = Convert.ToDouble(num);
                    double f = Math.Round(h, 0);
                    if ((f == 1) || (f == 2) || (f == 3) || (f == 4))
                    {
                        //Bugfix de cuando tiene enteros
                        double r = h;
                        r = r * 1000;
                        r = Math.Round(r, 0);
                        BusquedaF2(r.ToString());
                        //Busqueda hasta que encuentre un valor aproximado al que hay en la lista de referencia
                        double val2 = r;
                        while (check == 0)
                        {

                            val2 = val2 - 1;
                            //Busqueda al valor aproximado menor
                            BusquedaF2(val2.ToString());
                        }
                    }
                    else
                    {
                        BusquedaF2(f.ToString());
                        //Busqueda hasta que encuentre un valor aproximado al que hay en la lista de referencia
                        double val2 = f;
                        while (check == 0)
                        {

                            val2 = val2 - 1;
                            //Busqueda al valor aproximado menor
                            BusquedaF2(val2.ToString());
                        }
                    }
                    Row.Cells[3].Value = filaecu2;
                }
            }
            catch (Exception f)
            {

            }
            //Configuracion de los grid para que no admitan datos manuales
            dataGridView2.ReadOnly = true;
            dataGridView3.ReadOnly = true;
            dataGridView4.ReadOnly = true;
            dataGridView13.ReadOnly = true;
            dataGridView14.ReadOnly = true;
            dataGridView15.ReadOnly = true;

            valor_nominal.Add(Referencias.Rows[comboBox1.SelectedIndex].Cells[0].Value.ToString());

        }

        private void button3_Click(object sender, EventArgs e)
        {
            //Boton para borrar todos los registros y que se pueda volver a hacer una consulta con el mismo Excel pero con diferentes mallas seleccionadas
            dataGridView3.Visible = false;
            dataGridView4.Visible = false;

            dataGridView2.Rows.Clear();
            dataGridView3.Rows.Clear();
            dataGridView4.Rows.Clear();
            dataGridView5.Rows.Clear();
            dataGridView6.Rows.Clear();

            ch1 = true;
            ch2 = true;
            try
            {
                while (true)
                {
                    dataGridView5.Columns.RemoveAt(2);
                }
            }
            catch (Exception l)
            {

            }
            try
            {
                while (true)
                {
                    dataGridView3.Columns.RemoveAt(2);
                }
            }
            catch (Exception l)
            {

            }
            oc1 = true;
        }

        

        public void Busqueda(string num)
        {
            //Busqueda Pruebas para que me traiga la fila
            foreach (DataGridViewRow Row in dataGridView1.Rows)
            {
                strFila = Row.Index.ToString();
                Valor = Convert.ToString(Row.Cells[0].Value);
                if ((Valor == "Particle Size (µm)") || (Valor == ""))
                {

                }
                else
                {
                    double i = Convert.ToDouble(Valor);
                    i = i + .4;
                    double j = Math.Round(i, 0);
                    if (j.ToString() == num)
                    {
                        //Agregar esa fila a el datagrid "FilaSeleccionada"
                        FilaSeleccionada.Rows.Add(dataGridView1.Rows[Convert.ToInt32(strFila)].Cells[0].Value.ToString(), dataGridView1.Rows[Convert.ToInt32(strFila)].Cells[2].Value.ToString(), dataGridView1.Rows[Convert.ToInt32(strFila)].Cells[3].Value.ToString(), dataGridView1.Rows[Convert.ToInt32(strFila)].Cells[4].Value.ToString());
                        filaecu = strFila;
                        check = 1;
                    }
                }
            }
        }
        public void BusquedaT(string num)
        {
            //Busqueda Pruebas para que me traiga la fila
            foreach (DataGridViewRow Row in dataGridView1.Rows)
            {
                strFila = Row.Index.ToString();
                Valor = Convert.ToString(Row.Cells[0].Value);
                if ((Valor == "Particle Size (µm)") || (Valor == ""))
                {

                }
                else
                {
                    double i = Convert.ToDouble(Valor);
                    i = i + .4;
                    double j = Math.Round(i, 0);
                    if (j.ToString() == num)
                    {
                        //Agregar esa fila a el datagrid "FilaSeleccionada"
                        FilaSeleccionada.Rows.Add(dataGridView1.Rows[Convert.ToInt32(strFila)].Cells[0].Value.ToString(), dataGridView1.Rows[Convert.ToInt32(strFila)].Cells[2].Value.ToString(), dataGridView1.Rows[Convert.ToInt32(strFila)].Cells[3].Value.ToString(), dataGridView1.Rows[Convert.ToInt32(strFila)].Cells[4].Value.ToString());
                        filaecu = strFila;
                        check = 1;
                    }
                }
            }
        }

        public void BusquedaF2(string num)
        {
            //Busqueda Pruebas para que me traiga la fila
            foreach (DataGridViewRow Row in dataGridView1.Rows)
            {
                strFila2 = Row.Index.ToString();
                Valor2 = Convert.ToString(Row.Cells[0].Value);
                if ((Valor2 == "Particle Size (µm)") || (Valor2 == ""))
                {

                }
                else
                {
                    //Correccion para los puntos decimales que hacen que suba o no suba el valor
                    double i = Convert.ToDouble(Valor2);
                    i = i + .4;
                    double j = Math.Round(i, 0);
                    if (j.ToString() == num)
                    {
                        //Agregar esa fila a el datagrid "FilaSeleccionada"
                        FilaSeleccionada2.Rows.Add(dataGridView1.Rows[Convert.ToInt32(strFila2)].Cells[0].Value.ToString(), dataGridView1.Rows[Convert.ToInt32(strFila2)].Cells[2].Value.ToString(), dataGridView1.Rows[Convert.ToInt32(strFila2)].Cells[3].Value.ToString(), dataGridView1.Rows[Convert.ToInt32(strFila2)].Cells[4].Value.ToString());
                        filaecu2 = strFila2;
                        check = 1;
                    }
                }
            }
        }

        public void renombrar()
        {
            //Renombrar 95%
            //Se renombran las celdas 0 de los grids 4, 6
            foreach (DataGridViewRow renombre in dataGridView4.Rows)
            {
                try
                {
                    renombre.Cells[0].Value = dataGridView3.Rows[renombre.Index].Cells[1].Value;
                }
                catch (Exception re)
                {

                }
            }
            foreach (DataGridViewRow renombre1 in dataGridView6.Rows)
            {
                try
                {
                    renombre1.Cells[0].Value = dataGridView3.Rows[renombre1.Index].Cells[1].Value;
                }
                catch (Exception re)
                {

                }
            }
            //Renombrar max%
            //Se renombran las celdas 0 de los grids 12, 15
            foreach (DataGridViewRow renombre in dataGridView15.Rows)
            {
                try
                {
                    renombre.Cells[0].Value = dataGridView14.Rows[renombre.Index].Cells[1].Value;
                }
                catch (Exception re)
                {

                }
            }
            foreach (DataGridViewRow renombre1 in dataGridView12.Rows)
            {
                try
                {
                    renombre1.Cells[0].Value = dataGridView14.Rows[renombre1.Index].Cells[1].Value;
                }
                catch (Exception re)
                {

                }
            }
            try
            {
                //Para 95%
                dataGridView4.Rows[(dataGridView4.Rows.Count - 4)].Cells[0].Value = ("999");
                dataGridView4.Rows[(dataGridView4.Rows.Count - 3)].Cells[0].Value = (dataGridView3.Rows[(dataGridView3.Rows.Count - 2)].Cells[1].Value);
                dataGridView4.Rows[(dataGridView4.Rows.Count - 2)].Cells[0].Value = (dataGridView3.Rows[(dataGridView3.Rows.Count - 1)].Cells[1].Value);

                dataGridView6.Rows[(dataGridView4.Rows.Count - 4)].Cells[0].Value = ("999");
                dataGridView6.Rows[(dataGridView4.Rows.Count - 3)].Cells[0].Value = (dataGridView3.Rows[(dataGridView3.Rows.Count - 2)].Cells[1].Value);
                dataGridView6.Rows[(dataGridView4.Rows.Count - 2)].Cells[0].Value = (dataGridView3.Rows[(dataGridView3.Rows.Count - 1)].Cells[1].Value);
                
                //Para max%
                dataGridView15.Rows[(dataGridView15.Rows.Count - 4)].Cells[0].Value = ("999");
                dataGridView15.Rows[(dataGridView15.Rows.Count - 3)].Cells[0].Value = (dataGridView14.Rows[(dataGridView14.Rows.Count - 2)].Cells[1].Value);
                dataGridView15.Rows[(dataGridView15.Rows.Count - 2)].Cells[0].Value = (dataGridView14.Rows[(dataGridView14.Rows.Count - 1)].Cells[1].Value);

                dataGridView12.Rows[(dataGridView15.Rows.Count - 4)].Cells[0].Value = ("999");
                dataGridView12.Rows[(dataGridView15.Rows.Count - 3)].Cells[0].Value = (dataGridView14.Rows[(dataGridView14.Rows.Count - 2)].Cells[1].Value);
                dataGridView12.Rows[(dataGridView15.Rows.Count - 2)].Cells[0].Value = (dataGridView14.Rows[(dataGridView14.Rows.Count - 1)].Cells[1].Value);

            }
            catch (Exception df)
            {

            }
        }

        public void renombrar1()
        {
            //Para 95%
            //Se renombran las celdas 0 de los grids 4, 6
            foreach (DataGridViewRow renombre in dataGridView4.Rows)
            {
                try
                {
                    renombre.Cells[0].Value = dataGridView3.Rows[renombre.Index].Cells[1].Value;
                }
                catch (Exception re)
                {

                }
            }
            foreach (DataGridViewRow renombre1 in dataGridView6.Rows)
            {
                try
                {
                    renombre1.Cells[0].Value = dataGridView3.Rows[renombre1.Index].Cells[1].Value;
                }
                catch (Exception re)
                {

                }
            }
            //Para max%
            //Se renombran las celdas 0 de los grids 12, 5
            foreach (DataGridViewRow renombre in dataGridView15.Rows)
            {
                try
                {
                    renombre.Cells[0].Value = dataGridView14.Rows[renombre.Index].Cells[1].Value;
                }
                catch (Exception re)
                {

                }
            }
            foreach (DataGridViewRow renombre1 in dataGridView12.Rows)
            {
                try
                {
                    renombre1.Cells[0].Value = dataGridView14.Rows[renombre1.Index].Cells[1].Value;
                }
                catch (Exception re)
                {

                }
            }
            //Recorrido de los nombres hacia abajo dependiendo del numero de filas que haya 
            //Para 95%
            try
            {
                int rep = Convert.ToInt32(dataGridView4.RowCount) + 2;
                int cont = 2;
                int cont2 = 1;
                while (cont < rep)
                {
                    dataGridView4.Rows[dataGridView4.RowCount - cont2].Cells[0].Value = dataGridView4.Rows[dataGridView4.RowCount - cont].Cells[0].Value;
                    dataGridView6.Rows[dataGridView4.RowCount - cont2].Cells[0].Value = dataGridView4.Rows[dataGridView4.RowCount - cont].Cells[0].Value;

                    cont++;
                    cont2++;
                }
            }
            catch (Exception gh)
            {

            }

            //Para max%
            try
            {
                int rep1 = Convert.ToInt32(dataGridView15.RowCount) + 2;
                int cont1 = 2;
                int cont21 = 1;
                while (cont1 < rep1)
                {
                    dataGridView15.Rows[dataGridView15.RowCount - cont21].Cells[0].Value = dataGridView15.Rows[dataGridView15.RowCount - cont1].Cells[0].Value;
                    dataGridView12.Rows[dataGridView15.RowCount - cont21].Cells[0].Value = dataGridView15.Rows[dataGridView15.RowCount - cont1].Cells[0].Value;

                    cont1++;
                    cont21++;
                }
            }
            catch (Exception gh)
            {

            }
            //Valores por Default en las celdas 0 row 0
            dataGridView4.Rows[0].Cells[0].Value = "999";
            dataGridView6.Rows[0].Cells[0].Value = "999";
            dataGridView15.Rows[0].Cells[0].Value = "999";
            dataGridView12.Rows[0].Cells[0].Value = "999";
        }

        private void button6_Click(object sender, EventArgs e)
        {
            //regresa a la seleccion de las mallas para la realizacion del reporte
            allowSelect = true;
            tabControl1.SelectedTab = tabPage2;
            allowSelect = false;

            dataGridView5.Rows.Clear();
            dataGridView6.Rows.Clear();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            //Esconde los acumulativos a la derecha dependiendo del numero de corridas que se esten manejando
            con_ocu = "si";
            //Para 3 corridas
            if (num_corr == "3")
            {
                try
                {
                    foreach (DataGridViewRow row in dataGridView5.Rows)
                    {
                        foreach (DataGridViewColumn col in dataGridView5.Columns)
                        {
                            if (col.Index >= 5)
                            {
                                dataGridView5.Rows[row.Index].Cells[col.Index].Value = "";
                            }
                        }
                    }
                    foreach (DataGridViewRow row in dataGridView11.Rows)
                    {
                        foreach (DataGridViewColumn col in dataGridView11.Columns)
                        {
                            if (col.Index >= 5)
                            {
                                dataGridView11.Rows[row.Index].Cells[col.Index].Value = "";
                            }
                        }
                    }
                }
                catch (Exception df)
                {
                    MessageBox.Show("Columns have been removed");
                }
            }
            else if (num_corr == "2")
            {
                //Para 2 corridas
                try
                {
                    foreach (DataGridViewRow row in dataGridView5.Rows)
                    {
                        foreach (DataGridViewColumn col in dataGridView5.Columns)
                        {
                            if (col.Index >= 4)
                            {
                                dataGridView5.Rows[row.Index].Cells[col.Index].Value = "";
                            }
                        }
                    }
                    foreach (DataGridViewRow row in dataGridView11.Rows)
                    {
                        foreach (DataGridViewColumn col in dataGridView11.Columns)
                        {
                            if (col.Index >= 4)
                            {
                                dataGridView11.Rows[row.Index].Cells[col.Index].Value = "";
                            }
                        }
                    }
                }
                catch (Exception df)
                {
                    MessageBox.Show("Columns have been removed");
                }
            }
            else if (num_corr == "1")
            {
                //Para 1 corrida
                try
                {
                    foreach (DataGridViewRow row in dataGridView5.Rows)
                    {
                        foreach (DataGridViewColumn col in dataGridView5.Columns)
                        {
                            if (col.Index >= 3)
                            {
                                dataGridView5.Rows[row.Index].Cells[col.Index].Value = "";
                            }
                        }
                    }
                    foreach (DataGridViewRow row in dataGridView11.Rows)
                    {
                        foreach (DataGridViewColumn col in dataGridView11.Columns)
                        {
                            if (col.Index >= 3)
                            {
                                dataGridView11.Rows[row.Index].Cells[col.Index].Value = "";
                            }
                        }
                    }
                }
                catch (Exception df)
                {
                    MessageBox.Show("Columns have been removed");
                }
            }
            button7.Visible = false;
        }

        private void button10_Click(object sender, EventArgs e)
        {
            //Regresa a la vetana de Excel donde se puede seleccionar otro archivo excel para otro reporte
            Nombres.RemoveRange(0, Nombres.Count);
            Fecha.RemoveRange(0, Fecha.Count);
            Usuarios.RemoveRange(0, Usuarios.Count);
            Equipos.RemoveRange(0, Equipos.Count);
            Ids.RemoveRange(0, Ids.Count);
            Grupos.RemoveRange(0, Grupos.Count);
            Lotes.RemoveRange(0, Lotes.Count);
            Comentarios.RemoveRange(0, Comentarios.Count);
            Clientes.RemoveRange(0, Clientes.Count);

            allowSelect = true;
            tabControl1.SelectedTab = tabPage1;
            allowSelect = false;
        }

        private void tabControl1_Selecting(object sender, TabControlCancelEventArgs e)
        {
            //Para que el usuario no pueda cambiar el tabcontrol a placer y solo sea por medio de los botones
            if(!allowSelect) e.Cancel = true;
        }

        private void button12_Click(object sender, EventArgs e)
        {
            //Genera el reporte mandando a llamar a la funcion de finalizar
            Finalizar();
        }

        //Funciones para que los textbox solo acepten numeros
        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsLetter(e.KeyChar)) 
            {
                e.Handled = true; 
            }
            else if (Char.IsControl(e.KeyChar)) 
            {
                //e.Handled = false; 
            }
            else 
            {
                e.Handled = false; 
            }
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsLetter(e.KeyChar)) 
            {
                e.Handled = true; 
            }
            else if (Char.IsControl(e.KeyChar)) 
            {
                //e.Handled = false; 
            }
            else 
            {
                e.Handled = false; 
            }
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsLetter(e.KeyChar)) 
            {
                e.Handled = true; 
            }
            else if (Char.IsControl(e.KeyChar)) 
            {
                //e.Handled = false; 
            }
            else 
            {
                e.Handled = false; 
            }
        }

        public void Busqueda2(string num)
        {
            //Busqueda Pruebas para que me traiga la fila
            foreach (DataGridViewRow Row in dataGridView1.Rows)
            {
                strFila = Row.Index.ToString();
                Valor = Convert.ToString(Row.Cells[0].Value);
                if ((Valor == "Particle Size (µm)") || (Valor == ""))
                {

                }
                else
                {
                    double i = Convert.ToDouble(Valor);
                    i = i + .4;
                    double j = Math.Round(i, 0);
                    if (j.ToString() == num)
                    {
                        //Agregar esa fila a el datagrid "FilaSeleccionada"
                        filaecu = strFila;
                        check = 1;
                    }
                }
            }
        }
        public void Busqueda3(string num)
        {
            //Busqueda Pruebas para que me traiga la fila
            foreach (DataGridViewRow Row in dataGridView1.Rows)
            {
                strFila = Row.Index.ToString();
                Valor = Convert.ToString(Row.Cells[0].Value);
                if ((Valor == "Particle Size (µm)") || (Valor == ""))
                {

                }
                else
                {
                    double i = Convert.ToDouble(Valor);
                    double j = Math.Round(i, 1);
                    if (j.ToString() == num)
                    {
                        //Agregar esa fila a el datagrid "FilaSeleccionada"
                        filaecu = strFila;
                        check = 1;
                    }
                }
            }
        }

        //Aqui busco el que manda datos al reporte para modificarlo
        public void Finalizar()
        {
            if (label88.Visible == false)
            {
                label85.Text = "";
            }
            if (label89.Visible == false)
            {
                label86.Text = "";
            }
            if (label90.Visible == false)
            {
                label87.Text = "";
            }
            //Datos de Acumulativos
            DataSet1 ds = new DataSet1();
            DataTable dt = new DataTable();

            dt = ds.Tables["Datos_Reporte"];

            while (dataGridView5.Rows.Count != valor_nominal.Count)
            {
                valor_nominal.Add("");
            }
            //If condicion de cuantas corridas son
            if (num_corr == "3")
            {
                //dgv5 = 8 columnas
                //Lectura de todos los datos para generar el reporte
                for (int i = 0; i < (dataGridView5.Rows.Count); i++)
                {

                    dt.Rows.Add(
                        dataGridView5.Rows[i].Cells[0].Value,
                        dataGridView5.Rows[i].Cells[1].Value,
                        dataGridView5.Rows[i].Cells[2].Value,
                        dataGridView5.Rows[i].Cells[3].Value,
                        dataGridView5.Rows[i].Cells[4].Value,
                        dataGridView5.Rows[i].Cells[5].Value,
                        dataGridView5.Rows[i].Cells[6].Value,
                        dataGridView5.Rows[i].Cells[7].Value,

                        dataGridView6.Rows[i].Cells[1].Value,
                        dataGridView6.Rows[i].Cells[2].Value,
                        dataGridView6.Rows[i].Cells[3].Value,

                        (label13.Text),
                        (label4.Text + label14.Text),
                        (label5.Text + label15.Text),
                        (label6.Text + label16.Text),
                        (label7.Text + label17.Text),
                        (label8.Text + label18.Text),
                        (label9.Text + label19.Text),
                        (label10.Text + label20.Text),
                        (label11.Text + label21.Text),
                        (label12.Text + label22.Text),

                        (label59.Text),
                        (label40.Text + label31.Text),
                        (label39.Text + label30.Text),
                        (label38.Text + label29.Text),
                        (label37.Text + label28.Text),
                        (label36.Text + label27.Text),
                        (label35.Text + label26.Text),
                        (label34.Text + label25.Text),
                        (label33.Text + label24.Text),
                        (label32.Text + label23.Text),

                        (label60.Text),
                        (label58.Text + label49.Text),
                        (label57.Text + label48.Text),
                        (label56.Text + label47.Text),
                        (label55.Text + label46.Text),
                        (label54.Text + label45.Text),
                        (label53.Text + label44.Text),
                        (label52.Text + label43.Text),
                        (label51.Text + label42.Text),
                        (label50.Text + label41.Text),

                        dataGridView11.Rows[i].Cells[2].Value,
                        dataGridView11.Rows[i].Cells[3].Value,
                        dataGridView11.Rows[i].Cells[4].Value,
                        dataGridView11.Rows[i].Cells[5].Value,
                        dataGridView11.Rows[i].Cells[6].Value,
                        dataGridView11.Rows[i].Cells[7].Value,

                        dataGridView12.Rows[i].Cells[1].Value,
                        dataGridView12.Rows[i].Cells[2].Value,
                        dataGridView12.Rows[i].Cells[3].Value,
                        dataGridView11.Rows[i].Cells[0].Value, 
                        num_corr,
                        valor_nominal[i]
                        );
                }
                Vista_i vi = new Vista_i(dt);
                vi.Show();
            }
            else if (num_corr == "2")
            {
                //dgv5 = 6 columnas
                //Lectura de todos los datos para generar el reporte
                for (int i = 0; i < (dataGridView5.Rows.Count); i++)
                {
                    dt.Rows.Add(
                        dataGridView5.Rows[i].Cells[0].Value,
                        dataGridView5.Rows[i].Cells[1].Value,
                        dataGridView5.Rows[i].Cells[2].Value,
                        dataGridView5.Rows[i].Cells[3].Value,
                        dataGridView5.Rows[i].Cells[4].Value,
                        dataGridView5.Rows[i].Cells[5].Value,
                        "",
                        "",

                        dataGridView6.Rows[i].Cells[1].Value,
                        dataGridView6.Rows[i].Cells[2].Value,
                        "",

                        (label13.Text),
                        (label4.Text + label14.Text),
                        (label5.Text + label15.Text),
                        (label6.Text + label16.Text),
                        (label7.Text + label17.Text),
                        (label8.Text + label18.Text),
                        (label9.Text + label19.Text),
                        (label10.Text + label20.Text),
                        (label11.Text + label21.Text),
                        (label12.Text + label22.Text),

                        (label59.Text),
                        (label40.Text + label31.Text),
                        (label39.Text + label30.Text),
                        (label38.Text + label29.Text),
                        (label37.Text + label28.Text),
                        (label36.Text + label27.Text),
                        (label35.Text + label26.Text),
                        (label34.Text + label25.Text),
                        (label33.Text + label24.Text),
                        (label32.Text + label23.Text),

                        (label60.Text),
                        (label58.Text + label49.Text),
                        (label57.Text + label48.Text),
                        (label56.Text + label47.Text),
                        (label55.Text + label46.Text),
                        (label54.Text + label45.Text),
                        (label53.Text + label44.Text),
                        (label52.Text + label43.Text),
                        (label51.Text + label42.Text),
                        (label50.Text + label41.Text),

                        dataGridView11.Rows[i].Cells[2].Value,
                        dataGridView11.Rows[i].Cells[3].Value,
                        dataGridView11.Rows[i].Cells[4].Value,
                        dataGridView11.Rows[i].Cells[5].Value,
                        "",
                        "",

                        dataGridView12.Rows[i].Cells[1].Value,
                        dataGridView12.Rows[i].Cells[2].Value,
                        "",
                        dataGridView11.Rows[i].Cells[0].Value, 
                        num_corr,
                        valor_nominal[i]
                        );
                }
                Vista_i vi = new Vista_i(dt);
                vi.Show();
            }
            else if (num_corr == "1")
            {
                //dgv5 = 4 columnas
                //Lectura de todos los datos para generar el reporte
                for (int i = 0; i < (dataGridView5.Rows.Count); i++)
                {

                    dt.Rows.Add(
                        dataGridView5.Rows[i].Cells[0].Value,
                        dataGridView5.Rows[i].Cells[1].Value,
                        dataGridView5.Rows[i].Cells[2].Value,
                        dataGridView5.Rows[i].Cells[3].Value,
                        "",
                        "",
                        "",
                        "",

                        dataGridView6.Rows[i].Cells[1].Value,
                        "",
                        "",

                        (label13.Text),
                        (label4.Text + label14.Text),
                        (label5.Text + label15.Text),
                        (label6.Text + label16.Text),
                        (label7.Text + label17.Text),
                        (label8.Text + label18.Text),
                        (label9.Text + label19.Text),
                        (label10.Text + label20.Text),
                        (label11.Text + label21.Text),
                        (label12.Text + label22.Text),

                        (label59.Text),
                        (label40.Text + label31.Text),
                        (label39.Text + label30.Text),
                        (label38.Text + label29.Text),
                        (label37.Text + label28.Text),
                        (label36.Text + label27.Text),
                        (label35.Text + label26.Text),
                        (label34.Text + label25.Text),
                        (label33.Text + label24.Text),
                        (label32.Text + label23.Text),

                        (label60.Text),
                        (label58.Text + label49.Text),
                        (label57.Text + label48.Text),
                        (label56.Text + label47.Text),
                        (label55.Text + label46.Text),
                        (label54.Text + label45.Text),
                        (label53.Text + label44.Text),
                        (label52.Text + label43.Text),
                        (label51.Text + label42.Text),
                        (label50.Text + label41.Text),

                        dataGridView11.Rows[i].Cells[2].Value,
                        dataGridView11.Rows[i].Cells[3].Value,
                        "",
                        "",
                        "",
                        "",

                        dataGridView12.Rows[i].Cells[1].Value,
                        "",
                        "",
                        dataGridView11.Rows[i].Cells[0].Value, 
                        num_corr,
                        valor_nominal[i]
                        );
                }
                Vista_i vi = new Vista_i(dt);
                vi.Show();
            }
        }

        public void Busqueda4(string num)
        {
            //Busqueda Pruebas para que me traiga la fila
            foreach (DataGridViewRow Row in dataGridView1.Rows)
            {
                strFila = Row.Index.ToString();
                Valor = Convert.ToString(Row.Cells[0].Value);
                if ((Valor == "Particle Size (µm)") || (Valor == ""))
                {

                }
                else
                {
                    double i = Convert.ToDouble(Valor);
                    double j = Math.Ceiling(i);
                    if (j.ToString() == num)
                    {
                        //Agregar esa fila a el datagrid "FilaSeleccionada"
                        FilaSeleccionada.Rows.Add(dataGridView1.Rows[Convert.ToInt32(strFila)].Cells[0].Value.ToString(), dataGridView1.Rows[Convert.ToInt32(strFila)].Cells[2].Value.ToString(), dataGridView1.Rows[Convert.ToInt32(strFila)].Cells[3].Value.ToString(), dataGridView1.Rows[Convert.ToInt32(strFila)].Cells[4].Value.ToString());
                        filaecu = strFila;
                        check = 1;
                    }
                }
            }
        }

        private void button19_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows[0].Cells[4].Value.ToString() == "Run_3 (Vol%)")
            {
                //3 Corridas
                num_corr = "3";
                //Añadir los campos de "Acumulativos <"
                DataGridViewTextBoxColumn acu1 = new DataGridViewTextBoxColumn();
                acu1.HeaderText = "Run_1 Cumulative <";
                acu1.Width = 80;

                DataGridViewTextBoxColumn acu2 = new DataGridViewTextBoxColumn();
                acu2.HeaderText = "Run_2 Cumulative <";
                acu2.Width = 80;

                DataGridViewTextBoxColumn acu3 = new DataGridViewTextBoxColumn();
                acu3.HeaderText = "Run_3 Cumulative <";
                acu3.Width = 80;

                dataGridView3.Columns.Add(acu1);
                dataGridView3.Columns.Add(acu2);
                dataGridView3.Columns.Add(acu3);

                //para dgv 14
                DataGridViewTextBoxColumn acu1z = new DataGridViewTextBoxColumn();
                acu1z.HeaderText = "Run_1 Cumulative <";
                acu1z.Width = 80;

                DataGridViewTextBoxColumn acu2z = new DataGridViewTextBoxColumn();
                acu2z.HeaderText = "Run_2 Cumulative <";
                acu2z.Width = 80;

                DataGridViewTextBoxColumn acu3z = new DataGridViewTextBoxColumn();
                acu3z.HeaderText = "Run_3 Cumulative <";
                acu3z.Width = 80;

                dataGridView14.Columns.Add(acu1z);
                dataGridView14.Columns.Add(acu2z);
                dataGridView14.Columns.Add(acu3z);

                DataGridViewTextBoxColumn acu11 = new DataGridViewTextBoxColumn();
                acu11.HeaderText = "Run_1 Cumulative <";
                acu11.Width = 80;

                DataGridViewTextBoxColumn acu21 = new DataGridViewTextBoxColumn();
                acu21.HeaderText = "Run_2 Cumulative <";
                acu21.Width = 80;

                DataGridViewTextBoxColumn acu31 = new DataGridViewTextBoxColumn();
                acu31.HeaderText = "Run_3 Cumulative <";
                acu31.Width = 80;

                dataGridView5.Columns.Add(acu11);
                dataGridView5.Columns.Add(acu21);
                dataGridView5.Columns.Add(acu31);

                //MaxSA columnas acumulativas <
                DataGridViewTextBoxColumn acu111 = new DataGridViewTextBoxColumn();
                acu111.HeaderText = "Run_1 Cumulative <";
                acu111.Width = 80;

                DataGridViewTextBoxColumn acu211 = new DataGridViewTextBoxColumn();
                acu211.HeaderText = "Run_2 Cumulative <";
                acu211.Width = 80;

                DataGridViewTextBoxColumn acu311 = new DataGridViewTextBoxColumn();
                acu311.HeaderText = "Run_3 Cumulative <";
                acu311.Width = 80;

                dataGridView11.Columns.Add(acu111);
                dataGridView11.Columns.Add(acu211);
                dataGridView11.Columns.Add(acu311);

                //Aqui ira el calculo de la interpolacion de valores para 95%
                //primera corrida 
                foreach (DataGridViewRow row1 in dataGridView2.Rows)
                {

                    double acumarr = 0;
                    int n = 1;
                    //aumentar a la fila los valores acumulativos a la derecha (los que van arriba)
                    try
                    {
                        while (n <= Convert.ToInt32(row1.Cells[3].Value))
                        {
                            acumarr = acumarr + Convert.ToDouble(dataGridView1.Rows[n].Cells[2].Value);
                            n++;
                            if (acumarr > 100)
                            {
                                acumarr = 100;
                            }
                            dataGridView3.Rows[row1.Index].Cells[2].Value = Math.Round(acumarr, 2);
                            dataGridView5.Rows[row1.Index].Cells[2].Value = Math.Round(acumarr, 2);
                        }
                        //Asignacion de valores para interpolación
                        string x0 = Convert.ToDouble(dataGridView1.Rows[n - 1].Cells[0].Value).ToString();
                        string x1 = Convert.ToDouble(dataGridView1.Rows[n].Cells[0].Value).ToString();
                        string y0 = dataGridView5.Rows[row1.Index].Cells[2].Value.ToString();
                        acumarr = acumarr + Convert.ToDouble(dataGridView1.Rows[n].Cells[2].Value);
                        string y1 = Math.Round(acumarr, 2).ToString();

                        string texto = Convert.ToString(row1.Cells[2].Value);
                        Match m = Regex.Match(texto, "(\\d+)");
                        string num = string.Empty;
                        if (m.Success)
                        {
                            num = m.Value;
                        }
                        double h = Convert.ToDouble(num);

                        ////formula de interpolacion=
                        double arriba = Convert.ToDouble(h) - Convert.ToDouble(x0);
                        double abajo = Convert.ToDouble(x1) - Convert.ToDouble(x0);
                        double division = arriba / abajo;
                        double resultado = Convert.ToDouble(y0) + (division * (Convert.ToDouble(y1) - Convert.ToDouble(y0)));
                        dataGridView5.Rows[row1.Index].Cells[2].Value = Math.Round(resultado, 2);
                        dataGridView3.Rows[row1.Index].Cells[2].Value = Math.Round(resultado, 2);
                    }
                    catch (Exception r)
                    {

                    }
                }
                //Aqui ira el calculo de la interpolacion de valores para max%
                //primera corrida 
                foreach (DataGridViewRow row1 in dataGridView13.Rows)
                {

                    double acumarr = 0;
                    int n = 1;
                    //aumentar a la fila los valores acumulativos a la derecha (los que van arriba)
                    try
                    {
                        while (n <= Convert.ToInt32(row1.Cells[3].Value))
                        {
                            acumarr = acumarr + Convert.ToDouble(dataGridView1.Rows[n].Cells[2].Value);
                            n++;
                            if (acumarr > 100)
                            {
                                acumarr = 100;
                            }
                            dataGridView14.Rows[row1.Index].Cells[2].Value = Math.Round(acumarr, 2);
                            dataGridView11.Rows[row1.Index].Cells[2].Value = Math.Round(acumarr, 2);
                        }
                        //Asignacion de valores para interpolación
                        string x0 = Convert.ToDouble(dataGridView1.Rows[n - 1].Cells[0].Value).ToString();
                        string x1 = Convert.ToDouble(dataGridView1.Rows[n].Cells[0].Value).ToString();
                        string y0 = dataGridView11.Rows[row1.Index].Cells[2].Value.ToString();
                        acumarr = acumarr + Convert.ToDouble(dataGridView1.Rows[n].Cells[2].Value);
                        string y1 = Math.Round(acumarr, 2).ToString();

                        string texto = Convert.ToString(row1.Cells[2].Value);
                        Match m = Regex.Match(texto, "(\\d+)");
                        string num = string.Empty;
                        if (m.Success)
                        {
                            num = m.Value;
                        }
                        double h = Convert.ToDouble(num);

                        ////formula de interpolacion=
                        double arriba = Convert.ToDouble(h) - Convert.ToDouble(x0);
                        double abajo = Convert.ToDouble(x1) - Convert.ToDouble(x0);
                        double division = arriba / abajo;
                        double resultado = Convert.ToDouble(y0) + (division * (Convert.ToDouble(y1) - Convert.ToDouble(y0)));
                        dataGridView14.Rows[row1.Index].Cells[2].Value = Math.Round(resultado, 2);
                        dataGridView11.Rows[row1.Index].Cells[2].Value = Math.Round(resultado, 2);
                    }
                    catch (Exception r)
                    {

                    }
                }

                //segunda Corrida para 95%
                foreach (DataGridViewRow row2 in dataGridView2.Rows)
                {
                    double acumarr2 = 0;
                    int n2 = 1;
                    //aumentar a la fila los valores acumulativos a la derecha (los que van arriba)
                    try
                    {
                        while (n2 <= Convert.ToInt32(row2.Cells[3].Value))
                        {
                            acumarr2 = acumarr2 + Convert.ToDouble(dataGridView1.Rows[n2].Cells[3].Value);
                            n2++;
                            if (acumarr2 > 100)
                            {
                                acumarr2 = 100;
                            }
                            dataGridView3.Rows[row2.Index].Cells[3].Value = Math.Round(acumarr2, 2);
                            dataGridView5.Rows[row2.Index].Cells[3].Value = Math.Round(acumarr2, 2);
                        }
                        //Asignacion de valores para interpolación
                        string x0 = Convert.ToDouble(dataGridView1.Rows[n2 - 1].Cells[0].Value).ToString();
                        string x1 = Convert.ToDouble(dataGridView1.Rows[n2].Cells[0].Value).ToString();
                        string y0 = dataGridView5.Rows[row2.Index].Cells[3].Value.ToString();
                        acumarr2 = acumarr2 + Convert.ToDouble(dataGridView1.Rows[n2].Cells[3].Value);
                        string y1 = Math.Round(acumarr2, 2).ToString();

                        string texto = Convert.ToString(row2.Cells[2].Value);
                        Match m = Regex.Match(texto, "(\\d+)");
                        string num = string.Empty;
                        if (m.Success)
                        {
                            num = m.Value;
                        }
                        double h = Convert.ToDouble(num);

                        ////formula de interpolacion=
                        double arriba = Convert.ToDouble(h) - Convert.ToDouble(x0);
                        double abajo = Convert.ToDouble(x1) - Convert.ToDouble(x0);
                        double division = arriba / abajo;
                        double resultado = Convert.ToDouble(y0) + (division * (Convert.ToDouble(y1) - Convert.ToDouble(y0)));
                        dataGridView5.Rows[row2.Index].Cells[3].Value = Math.Round(resultado, 2);
                        dataGridView3.Rows[row2.Index].Cells[3].Value = Math.Round(resultado, 2);
                    }
                    catch (Exception r)
                    {

                    }
                }
                //segunda Corrida para max%
                foreach (DataGridViewRow row2 in dataGridView13.Rows)
                {
                    double acumarr2 = 0;
                    int n2 = 1;
                    //aumentar a la fila los valores acumulativos a la derecha (los que van arriba)
                    try
                    {
                        while (n2 <= Convert.ToInt32(row2.Cells[3].Value))
                        {
                            acumarr2 = acumarr2 + Convert.ToDouble(dataGridView1.Rows[n2].Cells[3].Value);
                            n2++;
                            if (acumarr2 > 100)
                            {
                                acumarr2 = 100;
                            }
                            dataGridView14.Rows[row2.Index].Cells[3].Value = Math.Round(acumarr2, 2);
                            dataGridView11.Rows[row2.Index].Cells[3].Value = Math.Round(acumarr2, 2);
                        }
                        //Asignacion de valores para interpolación
                        string x0 = Convert.ToDouble(dataGridView1.Rows[n2 - 1].Cells[0].Value).ToString();
                        string x1 = Convert.ToDouble(dataGridView1.Rows[n2].Cells[0].Value).ToString();
                        string y0 = dataGridView11.Rows[row2.Index].Cells[3].Value.ToString();
                        acumarr2 = acumarr2 + Convert.ToDouble(dataGridView1.Rows[n2].Cells[3].Value);
                        string y1 = Math.Round(acumarr2, 2).ToString();

                        string texto = Convert.ToString(row2.Cells[2].Value);
                        Match m = Regex.Match(texto, "(\\d+)");
                        string num = string.Empty;
                        if (m.Success)
                        {
                            num = m.Value;
                        }
                        double h = Convert.ToDouble(num);

                        ////formula de interpolacion=
                        double arriba = Convert.ToDouble(h) - Convert.ToDouble(x0);
                        double abajo = Convert.ToDouble(x1) - Convert.ToDouble(x0);
                        double division = arriba / abajo;
                        double resultado = Convert.ToDouble(y0) + (division * (Convert.ToDouble(y1) - Convert.ToDouble(y0)));
                        dataGridView14.Rows[row2.Index].Cells[3].Value = Math.Round(resultado, 2);
                        dataGridView11.Rows[row2.Index].Cells[3].Value = Math.Round(resultado, 2);
                    }
                    catch (Exception r)
                    {

                    }
                }

                //tercera Corrida para 95%
                foreach (DataGridViewRow row3 in dataGridView2.Rows)
                {
                    double acumarr3 = 0;
                    int n3 = 1;
                    //aumentar a la fila los valores acumulativos a la derecha (los que van arriba)
                    try
                    {
                        while (n3 <= Convert.ToInt32(row3.Cells[3].Value))
                        {
                            acumarr3 = acumarr3 + Convert.ToDouble(dataGridView1.Rows[n3].Cells[4].Value);
                            n3++;
                            if (acumarr3 > 100)
                            {
                                acumarr3 = 100;
                            }
                            dataGridView3.Rows[row3.Index].Cells[4].Value = Math.Round(acumarr3, 2);
                            dataGridView5.Rows[row3.Index].Cells[4].Value = Math.Round(acumarr3, 2);
                        }
                        //Asignacion de valores para interpolación
                        string x0 = Convert.ToDouble(dataGridView1.Rows[n3 - 1].Cells[0].Value).ToString();
                        string x1 = Convert.ToDouble(dataGridView1.Rows[n3].Cells[0].Value).ToString();
                        string y0 = dataGridView5.Rows[row3.Index].Cells[4].Value.ToString();
                        acumarr3 = acumarr3 + Convert.ToDouble(dataGridView1.Rows[n3].Cells[4].Value);
                        string y1 = Math.Round(acumarr3, 2).ToString();

                        string texto = Convert.ToString(row3.Cells[2].Value);
                        Match m = Regex.Match(texto, "(\\d+)");
                        string num = string.Empty;
                        if (m.Success)
                        {
                            num = m.Value;
                        }
                        double h = Convert.ToDouble(num);

                        ////formula de interpolacion=
                        double arriba = Convert.ToDouble(h) - Convert.ToDouble(x0);
                        double abajo = Convert.ToDouble(x1) - Convert.ToDouble(x0);
                        double division = arriba / abajo;
                        double resultado = Convert.ToDouble(y0) + (division * (Convert.ToDouble(y1) - Convert.ToDouble(y0)));
                        dataGridView5.Rows[row3.Index].Cells[4].Value = Math.Round(resultado, 2);
                        dataGridView3.Rows[row3.Index].Cells[4].Value = Math.Round(resultado, 2);
                    }
                    catch (Exception r)
                    {

                    }
                }
                //tercera Corrida para max%
                foreach (DataGridViewRow row3 in dataGridView13.Rows)
                {
                    double acumarr3 = 0;
                    int n3 = 1;
                    //aumentar a la fila los valores acumulativos a la derecha (los que van arriba)
                    try
                    {
                        while (n3 <= Convert.ToInt32(row3.Cells[3].Value))
                        {
                            acumarr3 = acumarr3 + Convert.ToDouble(dataGridView1.Rows[n3].Cells[4].Value);
                            n3++;
                            if (acumarr3 > 100)
                            {
                                acumarr3 = 100;
                            }
                            dataGridView14.Rows[row3.Index].Cells[4].Value = Math.Round(acumarr3, 2);
                            dataGridView11.Rows[row3.Index].Cells[4].Value = Math.Round(acumarr3, 2);
                        }
                        //Asignacion de valores para interpolación
                        string x0 = Convert.ToDouble(dataGridView1.Rows[n3 - 1].Cells[0].Value).ToString();
                        string x1 = Convert.ToDouble(dataGridView1.Rows[n3].Cells[0].Value).ToString();
                        string y0 = dataGridView11.Rows[row3.Index].Cells[4].Value.ToString();
                        acumarr3 = acumarr3 + Convert.ToDouble(dataGridView1.Rows[n3].Cells[4].Value);
                        string y1 = Math.Round(acumarr3, 2).ToString();

                        string texto = Convert.ToString(row3.Cells[2].Value);
                        Match m = Regex.Match(texto, "(\\d+)");
                        string num = string.Empty;
                        if (m.Success)
                        {
                            num = m.Value;
                        }
                        double h = Convert.ToDouble(num);

                        ////formula de interpolacion=
                        double arriba = Convert.ToDouble(h) - Convert.ToDouble(x0);
                        double abajo = Convert.ToDouble(x1) - Convert.ToDouble(x0);
                        double division = arriba / abajo;
                        double resultado = Convert.ToDouble(y0) + (division * (Convert.ToDouble(y1) - Convert.ToDouble(y0)));
                        dataGridView14.Rows[row3.Index].Cells[4].Value = Math.Round(resultado, 2);
                        dataGridView11.Rows[row3.Index].Cells[4].Value = Math.Round(resultado, 2);
                    }
                    catch (Exception r)
                    {

                    }
                }
                ch1 = false;
                dataGridView3.AllowUserToAddRows = false;
                dataGridView14.AllowUserToAddRows = false;

                //Otros acumulativos
                dataGridView3.Visible = true;
                dataGridView14.Visible = true;
                //Añadir los campos de "Acumulativos >"
                DataGridViewTextBoxColumn acu1z1 = new DataGridViewTextBoxColumn();
                acu1z1.HeaderText = "Run_1 Cumulative >";
                acu1z1.Width = 80;

                DataGridViewTextBoxColumn acu2z1 = new DataGridViewTextBoxColumn();
                acu2z1.HeaderText = "Run_2 Cumulative >";
                acu2z1.Width = 80;

                DataGridViewTextBoxColumn acu3z1 = new DataGridViewTextBoxColumn();
                acu3z1.HeaderText = "Run_3 Cumulative >";
                acu3z1.Width = 80;

                dataGridView3.Columns.Add(acu1z1);
                dataGridView3.Columns.Add(acu2z1);
                dataGridView3.Columns.Add(acu3z1);
                //Añadir los campos de "Acumulativos >" para dgv14 y dgv5 y dgv11
                DataGridViewTextBoxColumn acu1z1x = new DataGridViewTextBoxColumn();
                acu1z1x.HeaderText = "Run_1 Cumulative >";
                acu1z1x.Width = 80;

                DataGridViewTextBoxColumn acu2z1x = new DataGridViewTextBoxColumn();
                acu2z1x.HeaderText = "Run_2 Cumulative >";
                acu2z1x.Width = 80;

                DataGridViewTextBoxColumn acu3z1x = new DataGridViewTextBoxColumn();
                acu3z1x.HeaderText = "Run_3 Cumulative >";
                acu3z1x.Width = 80;

                dataGridView14.Columns.Add(acu1z1x);
                dataGridView14.Columns.Add(acu2z1x);
                dataGridView14.Columns.Add(acu3z1x);

                DataGridViewTextBoxColumn acu11z = new DataGridViewTextBoxColumn();
                acu11z.HeaderText = "Run_1 Cumulative >";
                acu11z.Width = 80;

                DataGridViewTextBoxColumn acu21z = new DataGridViewTextBoxColumn();
                acu21z.HeaderText = "Run_2 Cumulative >";
                acu21z.Width = 80;

                DataGridViewTextBoxColumn acu31z = new DataGridViewTextBoxColumn();
                acu31z.HeaderText = "Run_3 Cumulative >";
                acu31z.Width = 100;

                dataGridView5.Columns.Add(acu11z);
                dataGridView5.Columns.Add(acu21z);
                dataGridView5.Columns.Add(acu31z);

                DataGridViewTextBoxColumn acu11z1 = new DataGridViewTextBoxColumn();
                acu11z1.HeaderText = "Run_1 Cumulative >";
                acu11z1.Width = 80;

                DataGridViewTextBoxColumn acu21z1 = new DataGridViewTextBoxColumn();
                acu21z1.HeaderText = "Run_2 Cumulative >";
                acu21z1.Width = 80;

                DataGridViewTextBoxColumn acu31z1 = new DataGridViewTextBoxColumn();
                acu31z1.HeaderText = "Run_3 Cumulative >";
                acu31z1.Width = 100;

                dataGridView11.Columns.Add(acu11z1);
                dataGridView11.Columns.Add(acu21z1);
                dataGridView11.Columns.Add(acu31z1);

                //primera corrida 95%
                foreach (DataGridViewRow row2 in dataGridView2.Rows)
                {
                    double acumarr2 = 0;
                    int q = Convert.ToInt32(row2.Cells[3].Value) + 1;
                    //aumentar a la fila los valores acumulativos a la derecha (los que van arriba)
                    try
                    {
                        while (q > Convert.ToInt32(row2.Cells[3].Value))
                        {
                            acumarr2 = acumarr2 + Convert.ToDouble(dataGridView1.Rows[q].Cells[2].Value);
                            q++;
                            if (acumarr2 > 100)
                            {
                                acumarr2 = 100;
                            }
                            dataGridView3.Rows[row2.Index].Cells[5].Value = Math.Round(acumarr2, 2);
                            dataGridView5.Rows[row2.Index].Cells[5].Value = Math.Round(acumarr2, 2);
                        }
                        double valor = 100 - Convert.ToDouble(dataGridView5.Rows[row2.Index].Cells[2].Value);
                        dataGridView5.Rows[row2.Index].Cells[5].Value = Math.Round(valor, 2);
                        dataGridView3.Rows[row2.Index].Cells[5].Value = Math.Round(valor, 2);
                    }
                    catch (Exception r)
                    {

                    }
                }

                //primera corrida max%
                foreach (DataGridViewRow row1 in dataGridView13.Rows)
                {
                    double acumarr = 0;
                    int q = Convert.ToInt32(row1.Cells[3].Value) + 1;
                    //aumentar a la fila los valores acumulativos a la derecha (los que van arriba)
                    try
                    {
                        while (q > Convert.ToInt32(row1.Cells[3].Value))
                        {
                            acumarr = acumarr + Convert.ToDouble(dataGridView1.Rows[q].Cells[2].Value);
                            q++;
                            if (acumarr > 100)
                            {
                                acumarr = 100;
                            }
                            dataGridView14.Rows[row1.Index].Cells[5].Value = Math.Round(acumarr, 2);
                            dataGridView11.Rows[row1.Index].Cells[5].Value = Math.Round(acumarr, 2);
                        }
                        //Asignacion de valores para interpolación
                        string x0 = Convert.ToDouble(dataGridView1.Rows[q - 1].Cells[0].Value).ToString();
                        string x1 = Convert.ToDouble(dataGridView1.Rows[q].Cells[0].Value).ToString();
                        string y0 = dataGridView11.Rows[row1.Index].Cells[5].Value.ToString();
                        acumarr = acumarr + Convert.ToDouble(dataGridView1.Rows[q].Cells[2].Value);
                        string y1 = Math.Round(acumarr, 2).ToString();

                        string texto = Convert.ToString(row1.Cells[2].Value);
                        Match m = Regex.Match(texto, "(\\d+)");
                        string num = string.Empty;
                        if (m.Success)
                        {
                            num = m.Value;
                        }
                        double h = Convert.ToDouble(num);

                        ////formula de interpolacion=
                        double arriba = Convert.ToDouble(h) - Convert.ToDouble(x0);
                        double abajo = Convert.ToDouble(x1) - Convert.ToDouble(x0);
                        double division = arriba / abajo;
                        double resultado = Convert.ToDouble(y0) + (division * (Convert.ToDouble(y1) - Convert.ToDouble(y0)));
                        dataGridView14.Rows[row1.Index].Cells[5].Value = Math.Round(resultado, 2);
                        dataGridView11.Rows[row1.Index].Cells[5].Value = Math.Round(resultado, 2);
                    }
                    catch (Exception r)
                    {

                    }
                }

                //segunda Corrida 95%
                foreach (DataGridViewRow row2 in dataGridView2.Rows)
                {
                    double acumarr2 = 0;
                    int q1 = Convert.ToInt32(row2.Cells[3].Value) + 1;
                    //aumentar a la fila los valores acumulativos a la derecha (los que van arriba)
                    try
                    {
                        while (q1 > Convert.ToInt32(row2.Cells[3].Value))
                        {
                            acumarr2 = acumarr2 + Convert.ToDouble(dataGridView1.Rows[q1].Cells[3].Value);
                            q1++;
                            if (acumarr2 > 100)
                            {
                                acumarr2 = 100;
                            }
                            dataGridView3.Rows[row2.Index].Cells[6].Value = Math.Round(acumarr2, 2);
                            dataGridView5.Rows[row2.Index].Cells[6].Value = Math.Round(acumarr2, 2);
                        }
                    }
                    catch (Exception r)
                    {

                    }
                }
                //segunda Corrida max%
                foreach (DataGridViewRow row2 in dataGridView13.Rows)
                {
                    double acumarr2 = 0;
                    int q1 = Convert.ToInt32(row2.Cells[3].Value) + 1;
                    //aumentar a la fila los valores acumulativos a la derecha (los que van arriba)
                    try
                    {
                        while (q1 > Convert.ToInt32(row2.Cells[3].Value))
                        {
                            acumarr2 = acumarr2 + Convert.ToDouble(dataGridView1.Rows[q1].Cells[3].Value);
                            q1++;
                            if (acumarr2 > 100)
                            {
                                acumarr2 = 100;
                            }
                            dataGridView14.Rows[row2.Index].Cells[6].Value = Math.Round(acumarr2, 2);
                            dataGridView11.Rows[row2.Index].Cells[6].Value = Math.Round(acumarr2, 2);
                        }
                        //Asignacion de valores para interpolación
                        string x0 = Convert.ToDouble(dataGridView1.Rows[q1 - 1].Cells[0].Value).ToString();
                        string x1 = Convert.ToDouble(dataGridView1.Rows[q1].Cells[0].Value).ToString();
                        string y0 = dataGridView11.Rows[row2.Index].Cells[6].Value.ToString();
                        acumarr2 = acumarr2 + Convert.ToDouble(dataGridView1.Rows[q1].Cells[3].Value);
                        string y1 = Math.Round(acumarr2, 2).ToString();

                        string texto = Convert.ToString(row2.Cells[2].Value);
                        Match m = Regex.Match(texto, "(\\d+)");
                        string num = string.Empty;
                        if (m.Success)
                        {
                            num = m.Value;
                        }
                        double h = Convert.ToDouble(num);

                        ////formula de interpolacion=
                        double arriba = Convert.ToDouble(h) - Convert.ToDouble(x0);
                        double abajo = Convert.ToDouble(x1) - Convert.ToDouble(x0);
                        double division = arriba / abajo;
                        double resultado = Convert.ToDouble(y0) + (division * (Convert.ToDouble(y1) - Convert.ToDouble(y0)));
                        dataGridView14.Rows[row2.Index].Cells[6].Value = Math.Round(resultado, 2);
                        dataGridView11.Rows[row2.Index].Cells[6].Value = Math.Round(resultado, 2);
                    }
                    catch (Exception r)
                    {

                    }
                }

                //tercera Corrida 95%
                foreach (DataGridViewRow row3 in dataGridView2.Rows)
                {
                    double acumarr3 = 0;
                    int q3 = Convert.ToInt32(row3.Cells[3].Value) + 1;
                    //aumentar a la fila los valores acumulativos a la derecha (los que van arriba)
                    try
                    {
                        while (q3 > Convert.ToInt32(row3.Cells[3].Value))
                        {
                            acumarr3 = acumarr3 + Convert.ToDouble(dataGridView1.Rows[q3].Cells[4].Value);
                            q3++;
                            if (acumarr3 > 100)
                            {
                                acumarr3 = 100;
                            }
                            dataGridView3.Rows[row3.Index].Cells[7].Value = Math.Round(acumarr3, 2);
                            dataGridView5.Rows[row3.Index].Cells[7].Value = Math.Round(acumarr3, 2);
                        }
                        //Asignacion de valores para interpolación
                        string x0 = Convert.ToDouble(dataGridView1.Rows[q3 - 1].Cells[0].Value).ToString();
                        string x1 = Convert.ToDouble(dataGridView1.Rows[q3].Cells[0].Value).ToString();
                        string y0 = dataGridView5.Rows[row3.Index].Cells[7].Value.ToString();
                        acumarr3 = acumarr3 + Convert.ToDouble(dataGridView1.Rows[q3].Cells[4].Value);
                        string y1 = Math.Round(acumarr3, 2).ToString();

                        string texto = Convert.ToString(row3.Cells[2].Value);
                        Match m = Regex.Match(texto, "(\\d+)");
                        string num = string.Empty;
                        if (m.Success)
                        {
                            num = m.Value;
                        }
                        double h = Convert.ToDouble(num);

                        ////formula de interpolacion=
                        double arriba = Convert.ToDouble(h) - Convert.ToDouble(x0);
                        double abajo = Convert.ToDouble(x1) - Convert.ToDouble(x0);
                        double division = arriba / abajo;
                        double resultado = Convert.ToDouble(y0) + (division * (Convert.ToDouble(y1) - Convert.ToDouble(y0)));
                        dataGridView5.Rows[row3.Index].Cells[7].Value = Math.Round(resultado, 2);
                        dataGridView3.Rows[row3.Index].Cells[7].Value = Math.Round(resultado, 2);
                    }
                    catch (Exception r)
                    {

                    }
                }
                //tercera Corrida max%
                foreach (DataGridViewRow row3 in dataGridView13.Rows)
                {
                    double acumarr3 = 0;
                    int q3 = Convert.ToInt32(row3.Cells[3].Value) + 1;
                    //aumentar a la fila los valores acumulativos a la derecha (los que van arriba)
                    try
                    {
                        while (q3 > Convert.ToInt32(row3.Cells[3].Value))
                        {
                            acumarr3 = acumarr3 + Convert.ToDouble(dataGridView1.Rows[q3].Cells[4].Value);
                            q3++;
                            if (acumarr3 > 100)
                            {
                                acumarr3 = 100;
                            }
                            dataGridView14.Rows[row3.Index].Cells[7].Value = Math.Round(acumarr3, 2);
                            dataGridView11.Rows[row3.Index].Cells[7].Value = Math.Round(acumarr3, 2);
                        }
                        //Asignacion de valores para interpolación
                        string x0 = Convert.ToDouble(dataGridView1.Rows[q3 - 1].Cells[0].Value).ToString();
                        string x1 = Convert.ToDouble(dataGridView1.Rows[q3].Cells[0].Value).ToString();
                        string y0 = dataGridView11.Rows[row3.Index].Cells[7].Value.ToString();
                        acumarr3 = acumarr3 + Convert.ToDouble(dataGridView1.Rows[q3].Cells[4].Value);
                        string y1 = Math.Round(acumarr3, 2).ToString();

                        string texto = Convert.ToString(row3.Cells[2].Value);
                        Match m = Regex.Match(texto, "(\\d+)");
                        string num = string.Empty;
                        if (m.Success)
                        {
                            num = m.Value;
                        }
                        double h = Convert.ToDouble(num);

                        ////formula de interpolacion=
                        double arriba = Convert.ToDouble(h) - Convert.ToDouble(x0);
                        double abajo = Convert.ToDouble(x1) - Convert.ToDouble(x0);
                        double division = arriba / abajo;
                        double resultado = Convert.ToDouble(y0) + (division * (Convert.ToDouble(y1) - Convert.ToDouble(y0)));
                        dataGridView14.Rows[row3.Index].Cells[7].Value = Math.Round(resultado, 2);
                        dataGridView11.Rows[row3.Index].Cells[7].Value = Math.Round(resultado, 2);
                    }
                    catch (Exception r)
                    {

                    }
                }
                foreach (DataGridViewRow row in dataGridView3.Rows)
                    //Llenado de los acumulativos a la izquierda por medio de total a 100 
                {
                    double resultado = 100 - Convert.ToDouble(dataGridView5.Rows[row.Index].Cells[2].Value);
                    dataGridView5.Rows[row.Index].Cells[5].Value = Math.Round(resultado, 2);
                    dataGridView3.Rows[row.Index].Cells[5].Value = Math.Round(resultado, 2);

                    double resultado2 = 100 - Convert.ToDouble(dataGridView5.Rows[row.Index].Cells[3].Value);
                    dataGridView5.Rows[row.Index].Cells[6].Value = Math.Round(resultado2, 2);
                    dataGridView3.Rows[row.Index].Cells[6].Value = Math.Round(resultado2, 2);

                    double resultado3 = 100 - Convert.ToDouble(dataGridView5.Rows[row.Index].Cells[4].Value);
                    dataGridView5.Rows[row.Index].Cells[7].Value = Math.Round(resultado3, 2);
                    dataGridView3.Rows[row.Index].Cells[7].Value = Math.Round(resultado3, 2);
                }
                foreach (DataGridViewRow row in dataGridView14.Rows)
                {
                    double resultado = 100 - Convert.ToDouble(dataGridView11.Rows[row.Index].Cells[2].Value);
                    dataGridView11.Rows[row.Index].Cells[5].Value = Math.Round(resultado, 2);
                    dataGridView14.Rows[row.Index].Cells[5].Value = Math.Round(resultado, 2);

                    double resultado2 = 100 - Convert.ToDouble(dataGridView11.Rows[row.Index].Cells[3].Value);
                    dataGridView11.Rows[row.Index].Cells[6].Value = Math.Round(resultado2, 2);
                    dataGridView14.Rows[row.Index].Cells[6].Value = Math.Round(resultado2, 2);

                    double resultado3 = 100 - Convert.ToDouble(dataGridView11.Rows[row.Index].Cells[4].Value);
                    dataGridView11.Rows[row.Index].Cells[7].Value = Math.Round(resultado3, 2);
                    dataGridView14.Rows[row.Index].Cells[7].Value = Math.Round(resultado3, 2);
                }
            }
            else if (dataGridView1.Rows[0].Cells[3].Value.ToString() == "Run_2 (Vol%)")
            {
                //2 Corridas
                num_corr = "2";
                //Añadir los campos de "Acumulativos <" 95%
                DataGridViewTextBoxColumn acu1 = new DataGridViewTextBoxColumn();
                acu1.HeaderText = "Run_1 Cumulative <";
                acu1.Width = 80;

                DataGridViewTextBoxColumn acu2 = new DataGridViewTextBoxColumn();
                acu2.HeaderText = "Run_2 Cumulative <";
                acu2.Width = 80;

                dataGridView3.Columns.Add(acu1);
                dataGridView3.Columns.Add(acu2);

                DataGridViewTextBoxColumn acu11 = new DataGridViewTextBoxColumn();
                acu11.HeaderText = "Run_1 Cumulative <";
                acu11.Width = 80;

                DataGridViewTextBoxColumn acu21 = new DataGridViewTextBoxColumn();
                acu21.HeaderText = "Run_2 Cumulative <";
                acu21.Width = 80;

                dataGridView5.Columns.Add(acu11);
                dataGridView5.Columns.Add(acu21);
                //Añadir los campos de "Acumulativos <" max%
                DataGridViewTextBoxColumn acu1z = new DataGridViewTextBoxColumn();
                acu1z.HeaderText = "Run_1 Cumulative <";
                acu1z.Width = 80;

                DataGridViewTextBoxColumn acu2z = new DataGridViewTextBoxColumn();
                acu2z.HeaderText = "Run_2 Cumulative <";
                acu2z.Width = 80;

                dataGridView14.Columns.Add(acu1z);
                dataGridView14.Columns.Add(acu2z);

                DataGridViewTextBoxColumn acu11z = new DataGridViewTextBoxColumn();
                acu11z.HeaderText = "Run_1 Cumulative <";
                acu11z.Width = 80;

                DataGridViewTextBoxColumn acu21z = new DataGridViewTextBoxColumn();
                acu21z.HeaderText = "Run_2 Cumulative <";
                acu21z.Width = 80;

                dataGridView11.Columns.Add(acu11z);
                dataGridView11.Columns.Add(acu21z);

                //primera corrida 95%
                foreach (DataGridViewRow row1 in dataGridView2.Rows)
                {
                    //primera corrida
                    double acumarr = 0;
                    int n = 1;
                    //aumentar a la fila los valores acumulativos a la derecha (los que van arriba)
                    try
                    {
                        while (n <= Convert.ToInt32(row1.Cells[3].Value))
                        {
                            acumarr = acumarr + Convert.ToDouble(dataGridView1.Rows[n].Cells[2].Value);
                            n++;
                            if (acumarr > 100)
                            {
                                acumarr = 100;
                            }
                            dataGridView3.Rows[row1.Index].Cells[2].Value = Math.Round(acumarr, 2);
                            dataGridView5.Rows[row1.Index].Cells[2].Value = Math.Round(acumarr, 2);
                        }
                    }
                    catch (Exception r)
                    {

                    }
                }
                //primera corrida max%
                foreach (DataGridViewRow row1 in dataGridView13.Rows)
                {
                    //primera corrida
                    double acumarr = 0;
                    int n = 1;
                    //aumentar a la fila los valores acumulativos a la derecha (los que van arriba)
                    try
                    {
                        while (n <= Convert.ToInt32(row1.Cells[3].Value))
                        {
                            acumarr = acumarr + Convert.ToDouble(dataGridView1.Rows[n].Cells[2].Value);
                            n++;
                            if (acumarr > 100)
                            {
                                acumarr = 100;
                            }
                            dataGridView14.Rows[row1.Index].Cells[2].Value = Math.Round(acumarr, 2);
                            dataGridView11.Rows[row1.Index].Cells[2].Value = Math.Round(acumarr, 2);
                        }
                    }
                    catch (Exception r)
                    {

                    }
                }
                //segunda Corrida 95%
                foreach (DataGridViewRow row2 in dataGridView2.Rows)
                {
                    //primera corrida
                    double acumarr2 = 0;
                    int n2 = 1;
                    //aumentar a la fila los valores acumulativos a la derecha (los que van arriba)
                    try
                    {
                        while (n2 <= Convert.ToInt32(row2.Cells[3].Value))
                        {
                            acumarr2 = acumarr2 + Convert.ToDouble(dataGridView1.Rows[n2].Cells[3].Value);
                            n2++;
                            if (acumarr2 > 100)
                            {
                                acumarr2 = 100;
                            }
                            dataGridView3.Rows[row2.Index].Cells[3].Value = Math.Round(acumarr2, 2);
                            dataGridView5.Rows[row2.Index].Cells[3].Value = Math.Round(acumarr2, 2);
                        }
                    }
                    catch (Exception r)
                    {

                    }
                }
                //segunda Corrida max%
                foreach (DataGridViewRow row2 in dataGridView13.Rows)
                {
                    //primera corrida
                    double acumarr2 = 0;
                    int n2 = 1;
                    //aumentar a la fila los valores acumulativos a la derecha (los que van arriba)
                    try
                    {
                        while (n2 <= Convert.ToInt32(row2.Cells[3].Value))
                        {
                            acumarr2 = acumarr2 + Convert.ToDouble(dataGridView1.Rows[n2].Cells[3].Value);
                            n2++;
                            if (acumarr2 > 100)
                            {
                                acumarr2 = 100;
                            }
                            dataGridView14.Rows[row2.Index].Cells[3].Value = Math.Round(acumarr2, 2);
                            dataGridView11.Rows[row2.Index].Cells[3].Value = Math.Round(acumarr2, 2);
                        }
                    }
                    catch (Exception r)
                    {

                    }
                }
                ch1 = false;
                dataGridView3.AllowUserToAddRows = false;
                dataGridView14.AllowUserToAddRows = false;

                //Otros acumulativos
                dataGridView3.Visible = true;
                dataGridView14.Visible = true;
                //Añadir los campos de "Acumulativos >" 95%
                DataGridViewTextBoxColumn acu1z1 = new DataGridViewTextBoxColumn();
                acu1z1.HeaderText = "Run_1 Cumulative >";
                acu1z1.Width = 80;

                DataGridViewTextBoxColumn acu2z1 = new DataGridViewTextBoxColumn();
                acu2z1.HeaderText = "Run_2 Cumulative >";
                acu2z1.Width = 80;

                dataGridView3.Columns.Add(acu1z1);
                dataGridView3.Columns.Add(acu2z1);

                DataGridViewTextBoxColumn acu11z1 = new DataGridViewTextBoxColumn();
                acu11z1.HeaderText = "Run_1 Cumulative >";
                acu11z1.Width = 80;

                DataGridViewTextBoxColumn acu21z1 = new DataGridViewTextBoxColumn();
                acu21z1.HeaderText = "Run_2 Cumulative >";
                acu21z1.Width = 80;

                dataGridView5.Columns.Add(acu11z1);
                dataGridView5.Columns.Add(acu21z1);
                //Añadir los campos de "Acumulativos >" max%
                DataGridViewTextBoxColumn acu1z1x = new DataGridViewTextBoxColumn();
                acu1z1x.HeaderText = "Run_1 Cumulative >";
                acu1z1x.Width = 80;

                DataGridViewTextBoxColumn acu2z1x = new DataGridViewTextBoxColumn();
                acu2z1x.HeaderText = "Run_2 Cumulative >";
                acu2z1x.Width = 80;

                dataGridView14.Columns.Add(acu1z1x);
                dataGridView14.Columns.Add(acu2z1x);

                DataGridViewTextBoxColumn acu11z1x = new DataGridViewTextBoxColumn();
                acu11z1x.HeaderText = "Run_1 Cumulative >";
                acu11z1x.Width = 80;

                DataGridViewTextBoxColumn acu21z1x = new DataGridViewTextBoxColumn();
                acu21z1x.HeaderText = "Run_2 Cumulative >";
                acu21z1x.Width = 80;

                dataGridView11.Columns.Add(acu11z1x);
                dataGridView11.Columns.Add(acu21z1x);

                //primera corrida 95%
                foreach (DataGridViewRow row1 in dataGridView2.Rows)
                {
                    double acumarr = 0;
                    int q = Convert.ToInt32(row1.Cells[3].Value) + 1;
                    //aumentar a la fila los valores acumulativos a la derecha (los que van arriba)
                    try
                    {
                        while (q > Convert.ToInt32(row1.Cells[3].Value))
                        {
                            acumarr = acumarr + Convert.ToDouble(dataGridView1.Rows[q].Cells[2].Value);
                            q++;
                            if (acumarr > 100)
                            {
                                acumarr = 100;
                            }
                            dataGridView3.Rows[row1.Index].Cells[4].Value = Math.Round(acumarr, 2);
                            dataGridView5.Rows[row1.Index].Cells[4].Value = Math.Round(acumarr, 2);
                        }
                    }
                    catch (Exception r)
                    {

                    }
                }
                //primera corrida max%
                foreach (DataGridViewRow row1 in dataGridView13.Rows)
                {
                    double acumarr = 0;
                    int q = Convert.ToInt32(row1.Cells[3].Value) + 1;
                    //aumentar a la fila los valores acumulativos a la derecha (los que van arriba)
                    try
                    {
                        while (q > Convert.ToInt32(row1.Cells[3].Value))
                        {
                            acumarr = acumarr + Convert.ToDouble(dataGridView1.Rows[q].Cells[2].Value);
                            q++;
                            if (acumarr > 100)
                            {
                                acumarr = 100;
                            }
                            dataGridView14.Rows[row1.Index].Cells[4].Value = Math.Round(acumarr, 2);
                            dataGridView11.Rows[row1.Index].Cells[4].Value = Math.Round(acumarr, 2);
                        }
                    }
                    catch (Exception r)
                    {

                    }
                }

                //segunda Corrida 95%
                foreach (DataGridViewRow row2 in dataGridView2.Rows)
                {
                    double acumarr2 = 0;
                    int q1 = Convert.ToInt32(row2.Cells[3].Value) + 1;
                    //aumentar a la fila los valores acumulativos a la derecha (los que van arriba)
                    try
                    {
                        while (q1 > Convert.ToInt32(row2.Cells[3].Value))
                        {
                            acumarr2 = acumarr2 + Convert.ToDouble(dataGridView1.Rows[q1].Cells[3].Value);
                            q1++;
                            if (acumarr2 > 100)
                            {
                                acumarr2 = 100;
                            }
                            dataGridView3.Rows[row2.Index].Cells[5].Value = Math.Round(acumarr2, 2);
                            dataGridView5.Rows[row2.Index].Cells[5].Value = Math.Round(acumarr2, 2);
                        }
                    }
                    catch (Exception r)
                    {

                    }
                }
                //segunda Corrida max%
                foreach (DataGridViewRow row2 in dataGridView13.Rows)
                {
                    double acumarr2 = 0;
                    int q1 = Convert.ToInt32(row2.Cells[3].Value) + 1;
                    //aumentar a la fila los valores acumulativos a la derecha (los que van arriba)
                    try
                    {
                        while (q1 > Convert.ToInt32(row2.Cells[3].Value))
                        {
                            acumarr2 = acumarr2 + Convert.ToDouble(dataGridView1.Rows[q1].Cells[3].Value);
                            q1++;
                            if (acumarr2 > 100)
                            {
                                acumarr2 = 100;
                            }
                            dataGridView14.Rows[row2.Index].Cells[5].Value = Math.Round(acumarr2, 2);
                            dataGridView11.Rows[row2.Index].Cells[5].Value = Math.Round(acumarr2, 2);
                        }
                    }
                    catch (Exception r)
                    {

                    }
                }
                foreach (DataGridViewRow row in dataGridView3.Rows)
                //Llenado de los acumulativos a la izquierda por medio de total a 100 
                {
                    double resultado = 100 - Convert.ToDouble(dataGridView5.Rows[row.Index].Cells[2].Value);
                    dataGridView5.Rows[row.Index].Cells[4].Value = Math.Round(resultado, 2);

                    double resultado2 = 100 - Convert.ToDouble(dataGridView5.Rows[row.Index].Cells[3].Value);
                    dataGridView5.Rows[row.Index].Cells[5].Value = Math.Round(resultado2, 2);
                }
                foreach (DataGridViewRow row in dataGridView14.Rows)
                {
                    double resultado = 100 - Convert.ToDouble(dataGridView11.Rows[row.Index].Cells[2].Value);
                    dataGridView11.Rows[row.Index].Cells[4].Value = Math.Round(resultado, 2);

                    double resultado2 = 100 - Convert.ToDouble(dataGridView11.Rows[row.Index].Cells[3].Value);
                    dataGridView11.Rows[row.Index].Cells[5].Value = Math.Round(resultado2, 2);
                }
            }
            else if (dataGridView1.Rows[0].Cells[2].Value.ToString() == "Run_1 (Vol%)")
            {
                // 1 corrida
                num_corr = "1";
                //Añadir los campos de "Acumulativos <" 95%
                DataGridViewTextBoxColumn acu1 = new DataGridViewTextBoxColumn();
                acu1.HeaderText = "Run_1 Cumulative <";
                acu1.Width = 80;

                dataGridView3.Columns.Add(acu1);

                DataGridViewTextBoxColumn acu11 = new DataGridViewTextBoxColumn();
                acu11.HeaderText = "Run_1 Cumulative <";
                acu11.Width = 80;

                dataGridView5.Columns.Add(acu11);
                //Añadir los campos de "Acumulativos <" max%
                DataGridViewTextBoxColumn acu1z = new DataGridViewTextBoxColumn();
                acu1z.HeaderText = "Run_1 Cumulative <";
                acu1z.Width = 80;

                dataGridView14.Columns.Add(acu1z);

                DataGridViewTextBoxColumn acu11z = new DataGridViewTextBoxColumn();
                acu11z.HeaderText = "Run_1 Cumulative <";
                acu11z.Width = 80;

                dataGridView11.Columns.Add(acu11z);


                //primera corrida 95%
                foreach (DataGridViewRow row1 in dataGridView2.Rows)
                {
                    //primera corrida
                    double acumarr = 0;
                    int n = 1;
                    //aumentar a la fila los valores acumulativos a la derecha (los que van arriba)
                    try
                    {
                        while (n <= Convert.ToInt32(row1.Cells[3].Value))
                        {
                            acumarr = acumarr + Convert.ToDouble(dataGridView1.Rows[n].Cells[2].Value);
                            n++;
                            if (acumarr > 100)
                            {
                                acumarr = 100;
                            }
                            dataGridView3.Rows[row1.Index].Cells[2].Value = Math.Round(acumarr, 2);
                            dataGridView5.Rows[row1.Index].Cells[2].Value = Math.Round(acumarr, 2);
                        }
                    }
                    catch (Exception r)
                    {

                    }
                }
                //primera corrida max%
                foreach (DataGridViewRow row1 in dataGridView13.Rows)
                {
                    //primera corrida
                    double acumarr = 0;
                    int n = 1;
                    //aumentar a la fila los valores acumulativos a la derecha (los que van arriba)
                    try
                    {
                        while (n <= Convert.ToInt32(row1.Cells[3].Value))
                        {
                            acumarr = acumarr + Convert.ToDouble(dataGridView1.Rows[n].Cells[2].Value);
                            n++;
                            if (acumarr > 100)
                            {
                                acumarr = 100;
                            }
                            dataGridView14.Rows[row1.Index].Cells[2].Value = Math.Round(acumarr, 2);
                            dataGridView11.Rows[row1.Index].Cells[2].Value = Math.Round(acumarr, 2);
                        }
                    }
                    catch (Exception r)
                    {

                    }
                }

                ch1 = false;
                dataGridView3.AllowUserToAddRows = false;
                dataGridView14.AllowUserToAddRows = false;

                //Otros acumulativos
                dataGridView3.Visible = true;
                dataGridView14.Visible = true;
                //Añadir los campos de "Acumulativos >" 95%
                DataGridViewTextBoxColumn acu1z1 = new DataGridViewTextBoxColumn();
                acu1z1.HeaderText = "Run_1 Cumulative >";
                acu1z1.Width = 80;

                dataGridView3.Columns.Add(acu1z1);

                DataGridViewTextBoxColumn acu11z1 = new DataGridViewTextBoxColumn();
                acu11z1.HeaderText = "Run_1 Cumulative >";
                acu11z1.Width = 80;

                dataGridView5.Columns.Add(acu11z1);
                //Añadir los campos de "Acumulativos >" max%
                DataGridViewTextBoxColumn acu1z1x = new DataGridViewTextBoxColumn();
                acu1z1x.HeaderText = "Run_1 Cumulative >";
                acu1z1x.Width = 80;

                dataGridView14.Columns.Add(acu1z1x);

                DataGridViewTextBoxColumn acu11z1x = new DataGridViewTextBoxColumn();
                acu11z1x.HeaderText = "Run_1 Cumulative >";
                acu11z1x.Width = 80;

                dataGridView11.Columns.Add(acu11z1x);

                //primera corrida 95%
                foreach (DataGridViewRow row1 in dataGridView2.Rows)
                {
                    double acumarr = 0;
                    int q = Convert.ToInt32(row1.Cells[3].Value) + 1;
                    //aumentar a la fila los valores acumulativos a la derecha (los que van arriba)
                    try
                    {
                        while (q > Convert.ToInt32(row1.Cells[3].Value))
                        {
                            acumarr = acumarr + Convert.ToDouble(dataGridView1.Rows[q].Cells[2].Value);
                            q++;
                            if (acumarr > 100)
                            {
                                acumarr = 100;
                            }
                            dataGridView3.Rows[row1.Index].Cells[3].Value = Math.Round(acumarr, 2);
                            dataGridView5.Rows[row1.Index].Cells[3].Value = Math.Round(acumarr, 2);
                        }
                    }
                    catch (Exception r)
                    {

                    }
                }
                //primera corrida max%
                foreach (DataGridViewRow row1 in dataGridView13.Rows)
                {
                    double acumarr = 0;
                    int q = Convert.ToInt32(row1.Cells[3].Value) + 1;
                    //aumentar a la fila los valores acumulativos a la derecha (los que van arriba)
                    try
                    {
                        while (q > Convert.ToInt32(row1.Cells[3].Value))
                        {
                            acumarr = acumarr + Convert.ToDouble(dataGridView1.Rows[q].Cells[2].Value);
                            q++;
                            if (acumarr > 100)
                            {
                                acumarr = 100;
                            }
                            dataGridView14.Rows[row1.Index].Cells[3].Value = Math.Round(acumarr, 2);
                            dataGridView11.Rows[row1.Index].Cells[3].Value = Math.Round(acumarr, 2);
                        }
                    }
                    catch (Exception r)
                    {

                    }
                }
                corridas = 1;
                foreach (DataGridViewRow row in dataGridView3.Rows)
                {
                    double resultado = 100 - Convert.ToDouble(dataGridView5.Rows[row.Index].Cells[2].Value);
                    dataGridView5.Rows[row.Index].Cells[3].Value = Math.Round(resultado, 2);
                }
                foreach (DataGridViewRow row in dataGridView14.Rows)
                {
                    double resultado = 100 - Convert.ToDouble(dataGridView11.Rows[row.Index].Cells[2].Value);
                    dataGridView11.Rows[row.Index].Cells[3].Value = Math.Round(resultado, 2);
                }
            }

            dataGridView3.AllowUserToAddRows = false;
            dataGridView14.AllowUserToAddRows = false;
            ch2 = false;

            //Aqui empieza el proceso del diferencial
            try
            {
                if (dataGridView1.Rows[0].Cells[4].Value.ToString() == "Run_3 (Vol%)")
                {
                    // 3 corridas 
                    if (dataGridView3.Rows.Count == 2)
                    {
                        dataGridView4.Visible = true;
                        dataGridView15.Visible = true;
                        //Funciones del diferencial 95%
                        dataGridView4.Rows.Add();
                        dataGridView4.Rows.Add();
                        dataGridView4.Rows.Add();

                        dataGridView6.Rows.Add();
                        dataGridView6.Rows.Add();
                        dataGridView6.Rows.Add();
                        //Funciones del diferencial max%
                        dataGridView15.Rows.Add();
                        dataGridView15.Rows.Add();
                        dataGridView15.Rows.Add();

                        dataGridView12.Rows.Add();
                        dataGridView12.Rows.Add();
                        dataGridView12.Rows.Add();
                        //Asignacion de variables de comparacion 95%
                        double comp1;
                        double val1 = 100000;
                        foreach (DataGridViewRow max1 in dataGridView3.Rows)
                        {
                            string max11 = Convert.ToString(max1.Cells[2].Value);
                            comp1 = Convert.ToDouble(max11);
                            if (comp1 < val1)
                            {
                                val1 = comp1;
                                name = max1.Cells[0].Value.ToString();
                                dataGridView4.Rows[0].Cells[0].Value = name;
                                dataGridView6.Rows[0].Cells[0].Value = name;
                            }
                        }
                        dataGridView4.Rows[0].Cells[1].Value = Math.Round(val1, 2);
                        dataGridView6.Rows[0].Cells[1].Value = Math.Round(val1, 2);

                        //Asignacion de variables de comparacion max%
                        double comp1z;
                        double val1z = 100000;
                        foreach (DataGridViewRow max1 in dataGridView14.Rows)
                        {
                            string max11 = Convert.ToString(max1.Cells[2].Value);
                            comp1z = Convert.ToDouble(max11);
                            if (comp1z < val1z)
                            {
                                val1z = comp1z;
                                namez = max1.Cells[0].Value.ToString();
                                dataGridView15.Rows[0].Cells[0].Value = namez;
                                dataGridView12.Rows[0].Cells[0].Value = namez;
                            }
                        }
                        dataGridView15.Rows[0].Cells[1].Value = Math.Round(val1z, 2);
                        dataGridView12.Rows[0].Cells[1].Value = Math.Round(val1z, 2);

                        //2 95%
                        double comp2;
                        double val2 = 100000;
                        foreach (DataGridViewRow max2 in dataGridView3.Rows)
                        {
                            string max21 = Convert.ToString(max2.Cells[3].Value);
                            comp2 = Convert.ToDouble(max21);
                            if (comp2 < val2)
                            {
                                val2 = comp2;
                            }
                        }
                        dataGridView4.Rows[0].Cells[2].Value = Math.Round(val2, 2);
                        dataGridView6.Rows[0].Cells[2].Value = Math.Round(val2, 2);
                        //2 max%
                        double comp2z;
                        double val2z = 100000;
                        foreach (DataGridViewRow max2 in dataGridView14.Rows)
                        {
                            string max21 = Convert.ToString(max2.Cells[3].Value);
                            comp2z = Convert.ToDouble(max21);
                            if (comp2z < val2z)
                            {
                                val2z = comp2z;
                            }
                        }
                        dataGridView15.Rows[0].Cells[2].Value = Math.Round(val2z, 2);
                        dataGridView12.Rows[0].Cells[2].Value = Math.Round(val2z, 2);

                        //3 95%
                        double comp3;
                        double val3 = 100000;

                        foreach (DataGridViewRow max3 in dataGridView3.Rows)
                        {
                            string max31 = Convert.ToString(max3.Cells[4].Value);
                            comp3 = Convert.ToDouble(max31);
                            if (comp3 < val3)
                            {
                                val3 = comp3;
                            }
                        }
                        dataGridView4.Rows[0].Cells[3].Value = Math.Round(val3, 2);
                        dataGridView6.Rows[0].Cells[3].Value = Math.Round(val3, 2);
                        //3 max%
                        double comp3z;
                        double val3z = 100000;

                        foreach (DataGridViewRow max3 in dataGridView14.Rows)
                        {
                            string max31 = Convert.ToString(max3.Cells[4].Value);
                            comp3z = Convert.ToDouble(max31);
                            if (comp3z < val3z)
                            {
                                val3z = comp3z;
                            }
                        }
                        dataGridView15.Rows[0].Cells[3].Value = Math.Round(val3z, 2);
                        dataGridView12.Rows[0].Cells[3].Value = Math.Round(val3z, 2);

                        //4 95%
                        double comp4;
                        double val4 = 1000000;
                        string name2;
                        foreach (DataGridViewRow max4 in dataGridView3.Rows)
                        {
                            string max41 = Convert.ToString(max4.Cells[5].Value);
                            comp4 = Convert.ToDouble(max41);
                            if (comp4 < val4)
                            {
                                val4 = comp4;
                                name2 = max4.Cells[0].Value.ToString();
                                dataGridView4.Rows[2].Cells[0].Value = name2;
                                dataGridView6.Rows[2].Cells[0].Value = name2;
                            }
                        }
                        dataGridView4.Rows[2].Cells[1].Value = Math.Round(val4, 2);
                        dataGridView6.Rows[2].Cells[1].Value = Math.Round(val4, 2);
                        //4 max%
                        double comp4z;
                        double val4z = 1000000;
                        string name2z;
                        foreach (DataGridViewRow max4 in dataGridView14.Rows)
                        {
                            string max41 = Convert.ToString(max4.Cells[5].Value);
                            comp4z = Convert.ToDouble(max41);
                            if (comp4z < val4z)
                            {
                                val4z = comp4z;
                                name2z = max4.Cells[0].Value.ToString();
                                dataGridView15.Rows[2].Cells[0].Value = name2z;
                                dataGridView12.Rows[2].Cells[0].Value = name2z;
                            }
                        }
                        dataGridView15.Rows[2].Cells[1].Value = Math.Round(val4z, 2);
                        dataGridView12.Rows[2].Cells[1].Value = Math.Round(val4z, 2);

                        //5 95%
                        double comp5;
                        double val5 = 1000000;
                        foreach (DataGridViewRow max5 in dataGridView3.Rows)
                        {
                            string max51 = Convert.ToString(max5.Cells[6].Value);
                            comp5 = Convert.ToDouble(max51);
                            if (comp5 < val5)
                            {
                                val5 = comp5;
                            }
                        }
                        dataGridView4.Rows[2].Cells[2].Value = Math.Round(val5, 2);
                        dataGridView6.Rows[2].Cells[2].Value = Math.Round(val5, 2);
                        //5 max%
                        double comp5z;
                        double val5z = 1000000;
                        foreach (DataGridViewRow max5 in dataGridView14.Rows)
                        {
                            string max51 = Convert.ToString(max5.Cells[6].Value);
                            comp5z = Convert.ToDouble(max51);
                            if (comp5z < val5z)
                            {
                                val5z = comp5z;
                            }
                        }
                        dataGridView15.Rows[2].Cells[2].Value = Math.Round(val5z, 2);
                        dataGridView12.Rows[2].Cells[2].Value = Math.Round(val5z, 2);

                        //6 95%
                        double comp6;
                        double val6 = 10000;
                        foreach (DataGridViewRow max6 in dataGridView3.Rows)
                        {
                            string max61 = Convert.ToString(max6.Cells[7].Value);
                            comp6 = Convert.ToDouble(max61);
                            if (comp6 < val6)
                            {
                                val6 = comp6;
                            }
                        }
                        dataGridView4.Rows[2].Cells[3].Value = Math.Round(val6, 2);
                        dataGridView6.Rows[2].Cells[3].Value = Math.Round(val6, 2);
                        //6 max%
                        double comp6z;
                        double val6z = 10000;
                        foreach (DataGridViewRow max6 in dataGridView14.Rows)
                        {
                            string max61 = Convert.ToString(max6.Cells[7].Value);
                            comp6z = Convert.ToDouble(max61);
                            if (comp6z < val6z)
                            {
                                val6z = comp6z;
                            }
                        }
                        dataGridView15.Rows[2].Cells[3].Value = Math.Round(val6z, 2);
                        dataGridView12.Rows[2].Cells[3].Value = Math.Round(val6z, 2);

                        //Crear el diferencial 95%
                        dataGridView4.Rows[1].Cells[1].Value =
                        Math.Round(Convert.ToDouble(100 - (Convert.ToDouble(dataGridView4.Rows[2].Cells[1].Value) +
                        Convert.ToDouble(dataGridView4.Rows[0].Cells[1].Value))), 2);

                        dataGridView4.Rows[1].Cells[2].Value =
                        Math.Round(Convert.ToDouble(100 - (Convert.ToDouble(dataGridView4.Rows[2].Cells[2].Value) +
                        Convert.ToDouble(dataGridView4.Rows[0].Cells[2].Value))), 2);

                        dataGridView4.Rows[1].Cells[3].Value =
                        Math.Round(Convert.ToDouble(100 - (Convert.ToDouble(dataGridView4.Rows[2].Cells[3].Value) +
                        Convert.ToDouble(dataGridView4.Rows[0].Cells[3].Value))), 2);

                        dataGridView6.Rows[1].Cells[1].Value =
                        Math.Round(Convert.ToDouble(100 - (Convert.ToDouble(dataGridView4.Rows[2].Cells[1].Value) +
                        Convert.ToDouble(dataGridView4.Rows[0].Cells[1].Value))), 2);

                        dataGridView6.Rows[1].Cells[2].Value =
                        Math.Round(Convert.ToDouble(100 - (Convert.ToDouble(dataGridView4.Rows[2].Cells[2].Value) +
                        Convert.ToDouble(dataGridView4.Rows[0].Cells[2].Value))), 2);

                        dataGridView6.Rows[1].Cells[3].Value =
                        Math.Round(Convert.ToDouble(100 - (Convert.ToDouble(dataGridView4.Rows[2].Cells[3].Value) +
                        Convert.ToDouble(dataGridView4.Rows[0].Cells[3].Value))), 2);

                        //Crear el diferencial max%
                        dataGridView15.Rows[1].Cells[1].Value =
                        Math.Round(Convert.ToDouble(100 - (Convert.ToDouble(dataGridView15.Rows[2].Cells[1].Value) +
                        Convert.ToDouble(dataGridView15.Rows[0].Cells[1].Value))), 2);

                        dataGridView15.Rows[1].Cells[2].Value =
                        Math.Round(Convert.ToDouble(100 - (Convert.ToDouble(dataGridView15.Rows[2].Cells[2].Value) +
                        Convert.ToDouble(dataGridView15.Rows[0].Cells[2].Value))), 2);

                        dataGridView15.Rows[1].Cells[3].Value =
                        Math.Round(Convert.ToDouble(100 - (Convert.ToDouble(dataGridView15.Rows[2].Cells[3].Value) +
                        Convert.ToDouble(dataGridView15.Rows[0].Cells[3].Value))), 2);

                        dataGridView12.Rows[1].Cells[1].Value =
                        Math.Round(Convert.ToDouble(100 - (Convert.ToDouble(dataGridView15.Rows[2].Cells[1].Value) +
                        Convert.ToDouble(dataGridView15.Rows[0].Cells[1].Value))), 2);

                        dataGridView12.Rows[1].Cells[2].Value =
                        Math.Round(Convert.ToDouble(100 - (Convert.ToDouble(dataGridView15.Rows[2].Cells[2].Value) +
                        Convert.ToDouble(dataGridView15.Rows[0].Cells[2].Value))), 2);

                        dataGridView12.Rows[1].Cells[3].Value =
                        Math.Round(Convert.ToDouble(100 - (Convert.ToDouble(dataGridView15.Rows[2].Cells[3].Value) +
                        Convert.ToDouble(dataGridView15.Rows[0].Cells[3].Value))), 2);

                        renombrar();
                    }
                    else if (dataGridView3.Rows.Count > 2)
                    {
                        dataGridView4.Visible = true;
                        dataGridView15.Visible = true;

                        //Para 95%
                        int n_filas = Convert.ToInt32(dataGridView3.Rows.Count.ToString()) + 1;
                        int contador = 1;
                        while (contador < n_filas)
                        {
                            dataGridView4.Rows.Add();
                            dataGridView6.Rows.Add();
                            contador++;
                        }
                        //Para max%
                        int n_filas1 = Convert.ToInt32(dataGridView14.Rows.Count.ToString()) + 1;
                        int contador1 = 1;
                        while (contador1 < n_filas1)
                        {
                            dataGridView15.Rows.Add();
                            dataGridView12.Rows.Add();
                            contador1++;
                        }

                        //llenar la primera y la ultima columna 95%
                        double comp;
                        double val = 100000;
                        foreach (DataGridViewRow max in dataGridView3.Rows)
                        {
                            string max1 = Convert.ToString(max.Cells[2].Value);
                            comp = Convert.ToDouble(max1);
                            if (comp < val)
                            {
                                val = comp;
                                name = max.Cells[0].Value.ToString();
                                dataGridView4.Rows[0].Cells[0].Value = name;
                                dataGridView6.Rows[0].Cells[0].Value = name;
                            }
                        }
                        dataGridView4.Rows[0].Cells[1].Value = Math.Round(val, 2);
                        dataGridView6.Rows[0].Cells[1].Value = Math.Round(val, 2);

                        double comp1;
                        double val1 = 100000;
                        foreach (DataGridViewRow max1 in dataGridView3.Rows)
                        {
                            string max11 = Convert.ToString(max1.Cells[5].Value);
                            comp1 = Convert.ToDouble(max11);
                            if (comp1 < val1)
                            {
                                val1 = comp1;
                            }
                        }
                        dataGridView4.Rows[n_filas - 1].Cells[1].Value = Math.Round(val1, 2);
                        dataGridView6.Rows[n_filas - 1].Cells[1].Value = Math.Round(val1, 2);
                        //llenar la primera y la ultima columna max%
                        double comp1z;
                        double val1z = 100000;
                        foreach (DataGridViewRow max in dataGridView14.Rows)
                        {
                            string max1 = Convert.ToString(max.Cells[2].Value);
                            comp1z = Convert.ToDouble(max1);
                            if (comp1z < val1z)
                            {
                                val1z = comp1z;
                                namez = max.Cells[0].Value.ToString();
                                dataGridView15.Rows[0].Cells[0].Value = namez;
                                dataGridView12.Rows[0].Cells[0].Value = namez;
                            }
                        }
                        dataGridView15.Rows[0].Cells[1].Value = Math.Round(val1z, 2);
                        dataGridView12.Rows[0].Cells[1].Value = Math.Round(val1z, 2);

                        double comp1z1;
                        double val1z1 = 100000;
                        foreach (DataGridViewRow max1 in dataGridView14.Rows)
                        {
                            string max11 = Convert.ToString(max1.Cells[5].Value);
                            comp1z1 = Convert.ToDouble(max11);
                            if (comp1z1 < val1z1)
                            {
                                val1z1 = comp1z1;
                            }
                        }
                        dataGridView15.Rows[n_filas1 - 1].Cells[1].Value = Math.Round(val1z1, 2);
                        dataGridView12.Rows[n_filas1 - 1].Cells[1].Value = Math.Round(val1z1, 2);

                        //segunda carrera 95%
                        double comp2;
                        double val2 = 100000;
                        foreach (DataGridViewRow max21 in dataGridView3.Rows)
                        {
                            string max2 = Convert.ToString(max21.Cells[3].Value);
                            comp2 = Convert.ToDouble(max2);
                            if (comp2 < val2)
                            {
                                val2 = comp2;
                            }
                        }
                        dataGridView4.Rows[0].Cells[2].Value = Math.Round(val2, 2);
                        dataGridView6.Rows[0].Cells[2].Value = Math.Round(val2, 2);

                        double comp4;
                        double val4 = 100000;
                        foreach (DataGridViewRow max4 in dataGridView3.Rows)
                        {
                            string max41 = Convert.ToString(max4.Cells[6].Value);
                            comp4 = Convert.ToDouble(max41);
                            if (comp4 < val4)
                            {
                                val4 = comp4;
                                lname = max4.Cells[0].Value.ToString();
                                dataGridView4.Rows[n_filas - 1].Cells[0].Value = lname;
                                dataGridView6.Rows[n_filas - 1].Cells[0].Value = lname;
                            }
                        }
                        dataGridView4.Rows[n_filas - 1].Cells[2].Value = Math.Round(val4, 2);
                        dataGridView6.Rows[n_filas - 1].Cells[2].Value = Math.Round(val4, 2);
                        //segunda carrera max%
                        double comp2z;
                        double val2z = 100000;
                        foreach (DataGridViewRow max21 in dataGridView14.Rows)
                        {
                            string max2 = Convert.ToString(max21.Cells[3].Value);
                            comp2z = Convert.ToDouble(max2);
                            if (comp2z < val2z)
                            {
                                val2z = comp2z;
                            }
                        }
                        dataGridView15.Rows[0].Cells[2].Value = Math.Round(val2z, 2);
                        dataGridView12.Rows[0].Cells[2].Value = Math.Round(val2z, 2);

                        double comp4z;
                        double val4z = 100000;
                        foreach (DataGridViewRow max4 in dataGridView14.Rows)
                        {
                            string max41 = Convert.ToString(max4.Cells[6].Value);
                            comp4z = Convert.ToDouble(max41);
                            if (comp4z < val4z)
                            {
                                val4z = comp4z;
                                lnamez = max4.Cells[0].Value.ToString();
                                dataGridView15.Rows[n_filas1 - 1].Cells[0].Value = lnamez;
                                dataGridView12.Rows[n_filas1 - 1].Cells[0].Value = lnamez;
                            }
                        }
                        dataGridView15.Rows[n_filas1 - 1].Cells[2].Value = Math.Round(val4z, 2);
                        dataGridView12.Rows[n_filas1 - 1].Cells[2].Value = Math.Round(val4z, 2);

                        //tercera carrera 95%
                        double comp5;
                        double val5 = 100000;
                        foreach (DataGridViewRow max51 in dataGridView3.Rows)
                        {
                            string max5 = Convert.ToString(max51.Cells[4].Value);
                            comp5 = Convert.ToDouble(max5);
                            if (comp5 < val5)
                            {
                                val5 = comp5;
                            }
                        }
                        dataGridView4.Rows[0].Cells[3].Value = Math.Round(val5, 2);
                        dataGridView6.Rows[0].Cells[3].Value = Math.Round(val5, 2);

                        double comp6;
                        double val6 = 100000;
                        foreach (DataGridViewRow max6 in dataGridView3.Rows)
                        {
                            string max61 = Convert.ToString(max6.Cells[7].Value);
                            comp6 = Convert.ToDouble(max61);
                            if (comp6 < val6)
                            {
                                val6 = comp6;
                            }
                        }
                        dataGridView4.Rows[n_filas - 1].Cells[3].Value = Math.Round(val6, 2);
                        dataGridView6.Rows[n_filas - 1].Cells[3].Value = Math.Round(val6, 2);
                        //tercera carrera max%
                        double comp5z;
                        double val5z = 100000;
                        foreach (DataGridViewRow max51 in dataGridView14.Rows)
                        {
                            string max5 = Convert.ToString(max51.Cells[4].Value);
                            comp5z = Convert.ToDouble(max5);
                            if (comp5z < val5z)
                            {
                                val5z = comp5z;
                            }
                        }
                        dataGridView15.Rows[0].Cells[3].Value = Math.Round(val5z, 2);
                        dataGridView12.Rows[0].Cells[3].Value = Math.Round(val5z, 2);

                        double comp6z;
                        double val6z = 100000;
                        foreach (DataGridViewRow max6 in dataGridView14.Rows)
                        {
                            string max61 = Convert.ToString(max6.Cells[7].Value);
                            comp6z = Convert.ToDouble(max61);
                            if (comp6z < val6z)
                            {
                                val6z = comp6z;
                            }
                        }
                        dataGridView15.Rows[n_filas1 - 1].Cells[3].Value = Math.Round(val6z, 2);
                        dataGridView12.Rows[n_filas1 - 1].Cells[3].Value = Math.Round(val6z, 2);

                        //Para 95%
                        double acumulador = 0;
                        double acumulador1 = 0;
                        double acumulador2 = 0;
                        int com = 0;
                        int com1 = 0;
                        int com2 = 0;
                        double calculo;
                        double calculo1;
                        double calculo2;
                        int check = 1;
                        int check1 = 0;
                        //Para max%
                        double acumuladorz = 0;
                        double acumulador1z = 0;
                        double acumulador2z = 0;
                        int comz = 0;
                        int com1z = 0;
                        int com2z = 0;
                        double calculoz;
                        double calculo1z;
                        double calculo2z;
                        int checkz = 1;
                        int check1z = 0;

                        //Llenar los ultimos datos 95%
                        foreach (DataGridViewRow con in dataGridView3.Rows)
                        {
                            //Primera Corrida
                            while (com < 1)
                            {
                                acumulador = Convert.ToDouble(dataGridView4.Rows[0].Cells[1].Value);
                                com++;
                            }
                            if (con.Cells[0].Value.ToString() != name.ToString())
                            {
                                //Operaciones
                                calculo = Convert.ToDouble(con.Cells[2].Value) - Convert.ToDouble(dataGridView3.Rows[check1].Cells[2].Value.ToString());
                                dataGridView4.Rows[check].Cells[1].Value = Math.Round(calculo, 2);
                                dataGridView6.Rows[check].Cells[1].Value = Math.Round(calculo, 2);
                                //Aumento de acumulador
                                acumulador = acumulador + Convert.ToDouble(con.Cells[2].Value.ToString());
                            }
                            //Segunda Corrida
                            while (com1 < 1)
                            {
                                acumulador1 = Convert.ToDouble(dataGridView4.Rows[0].Cells[2].Value);
                                com1++;
                            }
                            if (con.Cells[0].Value.ToString() != name.ToString())
                            {
                                //Operaciones
                                calculo1 = ((Convert.ToDouble(con.Cells[3].Value)) - (Convert.ToDouble(dataGridView3.Rows[check1].Cells[3].Value.ToString())));
                                dataGridView4.Rows[check].Cells[2].Value = Math.Round(calculo1, 2);
                                dataGridView6.Rows[check].Cells[2].Value = Math.Round(calculo1, 2);

                                //Aumento de acumulador
                                acumulador1 = acumulador1 + Convert.ToDouble(con.Cells[3].Value.ToString());
                            }
                            //Tercera Corrida
                            while (com2 < 1)
                            {
                                acumulador2 = Convert.ToDouble(dataGridView4.Rows[0].Cells[3].Value);
                                com2++;
                            }
                            if (con.Cells[0].Value.ToString() != name.ToString())
                            {
                                //Operaciones
                                calculo2 = ((Convert.ToDouble(con.Cells[4].Value)) - (Convert.ToDouble(dataGridView3.Rows[check1].Cells[4].Value.ToString())));
                                dataGridView4.Rows[check].Cells[3].Value = Math.Round(calculo2, 2);
                                dataGridView6.Rows[check].Cells[3].Value = Math.Round(calculo2, 2);

                                //Aumento de acumulador
                                acumulador2 = acumulador2 + Convert.ToDouble(con.Cells[4].Value.ToString());
                                check++;
                                check1++;
                            }
                        }
                        //Llenar los ultimos datos max%
                        foreach (DataGridViewRow con in dataGridView14.Rows)
                        {
                            //Primera Corrida
                            while (comz < 1)
                            {
                                acumuladorz = Convert.ToDouble(dataGridView15.Rows[0].Cells[1].Value);
                                comz++;
                            }
                            if (con.Cells[0].Value.ToString() != namez.ToString())
                            {
                                //Operaciones
                                calculoz = Convert.ToDouble(con.Cells[2].Value) - Convert.ToDouble(dataGridView14.Rows[check1z].Cells[2].Value.ToString());
                                dataGridView15.Rows[checkz].Cells[1].Value = Math.Round(calculoz, 2);
                                dataGridView12.Rows[checkz].Cells[1].Value = Math.Round(calculoz, 2);
                                //Aumento de acumulador
                                acumuladorz = acumuladorz + Convert.ToDouble(con.Cells[2].Value.ToString());
                            }
                            //Segunda Corrida
                            while (com1z < 1)
                            {
                                acumulador1z = Convert.ToDouble(dataGridView15.Rows[0].Cells[2].Value);
                                com1z++;
                            }
                            if (con.Cells[0].Value.ToString() != namez.ToString())
                            {
                                //Operaciones
                                calculo1z = ((Convert.ToDouble(con.Cells[3].Value)) - (Convert.ToDouble(dataGridView14.Rows[check1z].Cells[3].Value.ToString())));
                                dataGridView15.Rows[checkz].Cells[2].Value = Math.Round(calculo1z, 2);
                                dataGridView12.Rows[checkz].Cells[2].Value = Math.Round(calculo1z, 2);

                                //Aumento de acumulador
                                acumulador1z = acumulador1z + Convert.ToDouble(con.Cells[3].Value.ToString());
                            }
                            //Tercera Corrida
                            while (com2z < 1)
                            {
                                acumulador2z = Convert.ToDouble(dataGridView15.Rows[0].Cells[3].Value);
                                com2z++;
                            }
                            if (con.Cells[0].Value.ToString() != namez.ToString())
                            {
                                //Operaciones
                                calculo2z = ((Convert.ToDouble(con.Cells[4].Value)) - (Convert.ToDouble(dataGridView14.Rows[check1z].Cells[4].Value.ToString())));
                                dataGridView15.Rows[checkz].Cells[3].Value = Math.Round(calculo2z, 2);
                                dataGridView12.Rows[checkz].Cells[3].Value = Math.Round(calculo2z, 2);

                                //Aumento de acumulador
                                acumulador2z = acumulador2z + Convert.ToDouble(con.Cells[4].Value.ToString());
                                checkz++;
                                check1z++;
                            }
                        }
                        renombrar1();
                    }
                }
                else if (dataGridView1.Rows[0].Cells[3].Value.ToString() == "Run_2 (Vol%)")
                {
                    //2 corridas
                    if (dataGridView3.Rows.Count == 2)
                    {
                        dataGridView4.Visible = true;
                        dataGridView15.Visible = true;
                        //Funciones del diferencial 95%
                        dataGridView4.Rows.Add();
                        dataGridView4.Rows.Add();
                        dataGridView4.Rows.Add();

                        dataGridView6.Rows.Add();
                        dataGridView6.Rows.Add();
                        dataGridView6.Rows.Add();
                        //Funciones del diferencial max%
                        dataGridView15.Rows.Add();
                        dataGridView15.Rows.Add();
                        dataGridView15.Rows.Add();

                        dataGridView12.Rows.Add();
                        dataGridView12.Rows.Add();
                        dataGridView12.Rows.Add();

                        //Asignacion de variables de comparacion 95%
                        double comp1;
                        double val1 = 100000;
                        foreach (DataGridViewRow max1 in dataGridView3.Rows)
                        {
                            string max11 = Convert.ToString(max1.Cells[2].Value);
                            comp1 = Convert.ToDouble(max11);
                            if (comp1 < val1)
                            {
                                val1 = comp1;
                                name = max1.Cells[0].Value.ToString();
                                dataGridView4.Rows[0].Cells[0].Value = name;
                                dataGridView6.Rows[0].Cells[0].Value = name;
                            }
                        }
                        dataGridView4.Rows[0].Cells[1].Value = Math.Round(val1, 2);
                        dataGridView6.Rows[0].Cells[1].Value = Math.Round(val1, 2);
                        //Asignacion de variables de comparacion max%
                        double comp1z;
                        double val1z = 100000;
                        foreach (DataGridViewRow max1 in dataGridView14.Rows)
                        {
                            string max11 = Convert.ToString(max1.Cells[2].Value);
                            comp1z = Convert.ToDouble(max11);
                            if (comp1z < val1z)
                            {
                                val1z = comp1z;
                                namez = max1.Cells[0].Value.ToString();
                                dataGridView15.Rows[0].Cells[0].Value = namez;
                                dataGridView12.Rows[0].Cells[0].Value = namez;
                            }
                        }
                        dataGridView15.Rows[0].Cells[1].Value = Math.Round(val1z, 2);
                        dataGridView12.Rows[0].Cells[1].Value = Math.Round(val1z, 2);

                        //2 95%
                        double comp2;
                        double val2 = 100000;
                        foreach (DataGridViewRow max2 in dataGridView3.Rows)
                        {
                            string max21 = Convert.ToString(max2.Cells[3].Value);
                            comp2 = Convert.ToDouble(max21);
                            if (comp2 < val2)
                            {
                                val2 = comp2;
                            }
                        }
                        dataGridView4.Rows[0].Cells[2].Value = Math.Round(val2, 2);
                        dataGridView6.Rows[0].Cells[2].Value = Math.Round(val2, 2);
                        //2 95%
                        double comp2z;
                        double val2z = 100000;
                        foreach (DataGridViewRow max2 in dataGridView14.Rows)
                        {
                            string max21 = Convert.ToString(max2.Cells[3].Value);
                            comp2z = Convert.ToDouble(max21);
                            if (comp2z < val2z)
                            {
                                val2z = comp2z;
                            }
                        }
                        dataGridView15.Rows[0].Cells[2].Value = Math.Round(val2z, 2);
                        dataGridView12.Rows[0].Cells[2].Value = Math.Round(val2z, 2);

                        //4 95%
                        double comp4;
                        double val4 = 1000000;
                        string name2;
                        foreach (DataGridViewRow max4 in dataGridView3.Rows)
                        {
                            string max41 = Convert.ToString(max4.Cells[4].Value);
                            comp4 = Convert.ToDouble(max41);
                            if (comp4 < val4)
                            {
                                val4 = comp4;
                                name2 = max4.Cells[0].Value.ToString();
                                dataGridView4.Rows[2].Cells[0].Value = name2;
                                dataGridView6.Rows[2].Cells[0].Value = name2;
                            }
                        }
                        dataGridView4.Rows[2].Cells[1].Value = Math.Round(val4, 2);
                        dataGridView6.Rows[2].Cells[1].Value = Math.Round(val4, 2);
                        //4 max%
                        double comp4z;
                        double val4z = 1000000;
                        string name2z;
                        foreach (DataGridViewRow max4 in dataGridView14.Rows)
                        {
                            string max41 = Convert.ToString(max4.Cells[4].Value);
                            comp4z = Convert.ToDouble(max41);
                            if (comp4z < val4z)
                            {
                                val4z = comp4z;
                                name2z = max4.Cells[0].Value.ToString();
                                dataGridView15.Rows[2].Cells[0].Value = name2z;
                                dataGridView12.Rows[2].Cells[0].Value = name2z;
                            }
                        }
                        dataGridView15.Rows[2].Cells[1].Value = Math.Round(val4z, 2);
                        dataGridView12.Rows[2].Cells[1].Value = Math.Round(val4z, 2);

                        //5 95%
                        double comp5;
                        double val5 = 1000000;
                        foreach (DataGridViewRow max5 in dataGridView3.Rows)
                        {
                            string max51 = Convert.ToString(max5.Cells[5].Value);
                            comp5 = Convert.ToDouble(max51);
                            if (comp5 < val5)
                            {
                                val5 = comp5;
                            }
                        }
                        dataGridView4.Rows[2].Cells[2].Value = Math.Round(val5, 2);
                        dataGridView6.Rows[2].Cells[2].Value = Math.Round(val5, 2);
                        //5 max%
                        double comp5z;
                        double val5z = 1000000;
                        foreach (DataGridViewRow max5 in dataGridView14.Rows)
                        {
                            string max51 = Convert.ToString(max5.Cells[5].Value);
                            comp5z = Convert.ToDouble(max51);
                            if (comp5z < val5z)
                            {
                                val5z = comp5z;
                            }
                        }
                        dataGridView15.Rows[2].Cells[2].Value = Math.Round(val5z, 2);
                        dataGridView12.Rows[2].Cells[2].Value = Math.Round(val5z, 2);

                        //Crear el diferencial 95%
                        dataGridView4.Rows[1].Cells[1].Value =
                        Math.Round(Convert.ToDouble(100 - (Convert.ToDouble(dataGridView4.Rows[2].Cells[1].Value) +
                        Convert.ToDouble(dataGridView4.Rows[0].Cells[1].Value))), 2);

                        dataGridView4.Rows[1].Cells[2].Value =
                        Math.Round(Convert.ToDouble(100 - (Convert.ToDouble(dataGridView4.Rows[2].Cells[2].Value) +
                        Convert.ToDouble(dataGridView4.Rows[0].Cells[2].Value))), 2);

                        dataGridView6.Rows[1].Cells[1].Value =
                        Math.Round(Convert.ToDouble(100 - (Convert.ToDouble(dataGridView4.Rows[2].Cells[1].Value) +
                        Convert.ToDouble(dataGridView4.Rows[0].Cells[1].Value))), 2);

                        dataGridView6.Rows[1].Cells[2].Value =
                        Math.Round(Convert.ToDouble(100 - (Convert.ToDouble(dataGridView4.Rows[2].Cells[2].Value) +
                        Convert.ToDouble(dataGridView4.Rows[0].Cells[2].Value))), 2);
        
                        //Crear el diferencial max%
                        dataGridView15.Rows[1].Cells[1].Value =
                        Math.Round(Convert.ToDouble(100 - (Convert.ToDouble(dataGridView15.Rows[2].Cells[1].Value) +
                        Convert.ToDouble(dataGridView15.Rows[0].Cells[1].Value))), 2);

                        dataGridView15.Rows[1].Cells[2].Value =
                        Math.Round(Convert.ToDouble(100 - (Convert.ToDouble(dataGridView15.Rows[2].Cells[2].Value) +
                        Convert.ToDouble(dataGridView15.Rows[0].Cells[2].Value))), 2);

                        dataGridView12.Rows[1].Cells[1].Value =
                        Math.Round(Convert.ToDouble(100 - (Convert.ToDouble(dataGridView15.Rows[2].Cells[1].Value) +
                        Convert.ToDouble(dataGridView15.Rows[0].Cells[1].Value))), 2);

                        dataGridView12.Rows[1].Cells[2].Value =
                        Math.Round(Convert.ToDouble(100 - (Convert.ToDouble(dataGridView15.Rows[2].Cells[2].Value) +
                        Convert.ToDouble(dataGridView15.Rows[0].Cells[2].Value))), 2);

                        renombrar();
                    }
                    else if (dataGridView3.Rows.Count > 2)
                    {
                        dataGridView4.Visible = true;
                        dataGridView15.Visible = true;

                        //95%
                        int n_filas = Convert.ToInt32(dataGridView3.Rows.Count.ToString()) + 1;
                        int contador = 1;
                        while (contador < n_filas)
                        {
                            dataGridView4.Rows.Add();
                            dataGridView6.Rows.Add();
                            contador++;
                        }
                        //max%
                        int n_filas1 = Convert.ToInt32(dataGridView14.Rows.Count.ToString()) + 1;
                        int contador1 = 1;
                        while (contador1 < n_filas1)
                        {
                            dataGridView15.Rows.Add();
                            dataGridView12.Rows.Add();
                            contador1++;
                        }

                        //llenar la primera y la ultima columna 95%
                        double comp;
                        double val = 100000;
                        foreach (DataGridViewRow max in dataGridView3.Rows)
                        {
                            string max1 = Convert.ToString(max.Cells[2].Value);
                            comp = Convert.ToDouble(max1);
                            if (comp < val)
                            {
                                val = comp;
                                name = max.Cells[0].Value.ToString();
                                dataGridView4.Rows[0].Cells[0].Value = name;
                                dataGridView6.Rows[0].Cells[0].Value = name;
                            }
                        }
                        dataGridView4.Rows[0].Cells[1].Value = Math.Round(val, 2);
                        dataGridView6.Rows[0].Cells[1].Value = Math.Round(val, 2);

                        double comp1;
                        double val1 = 100000;
                        foreach (DataGridViewRow max1 in dataGridView3.Rows)
                        {
                            string max11 = Convert.ToString(max1.Cells[4].Value);
                            comp1 = Convert.ToDouble(max11);
                            if (comp1 < val1)
                            {
                                val1 = comp1;
                            }
                        }
                        dataGridView4.Rows[n_filas - 1].Cells[1].Value = Math.Round(val1, 2);
                        dataGridView6.Rows[n_filas - 1].Cells[1].Value = Math.Round(val1, 2);
                        //llenar la primera y la ultima columna max%
                        double compz;
                        double valz = 100000;
                        foreach (DataGridViewRow max in dataGridView14.Rows)
                        {
                            string max1 = Convert.ToString(max.Cells[2].Value);
                            compz = Convert.ToDouble(max1);
                            if (compz < valz)
                            {
                                valz = compz;
                                namez = max.Cells[0].Value.ToString();
                                dataGridView15.Rows[0].Cells[0].Value = namez;
                                dataGridView12.Rows[0].Cells[0].Value = namez;
                            }
                        }
                        dataGridView15.Rows[0].Cells[1].Value = Math.Round(valz, 2);
                        dataGridView12.Rows[0].Cells[1].Value = Math.Round(valz, 2);

                        double comp1z;
                        double val1z = 100000;
                        foreach (DataGridViewRow max1 in dataGridView14.Rows)
                        {
                            string max11 = Convert.ToString(max1.Cells[4].Value);
                            comp1z = Convert.ToDouble(max11);
                            if (comp1z < val1z)
                            {
                                val1z = comp1z;
                            }
                        }
                        dataGridView15.Rows[n_filas1 - 1].Cells[1].Value = Math.Round(val1z, 2);
                        dataGridView12.Rows[n_filas1 - 1].Cells[1].Value = Math.Round(val1z, 2);

                        //segunda carrera 95%
                        double comp2;
                        double val2 = 100000;
                        foreach (DataGridViewRow max21 in dataGridView3.Rows)
                        {
                            string max2 = Convert.ToString(max21.Cells[3].Value);
                            comp2 = Convert.ToDouble(max2);
                            if (comp2 < val2)
                            {
                                val2 = comp2;
                            }
                        }
                        dataGridView4.Rows[0].Cells[2].Value = Math.Round(val2, 2);
                        dataGridView6.Rows[0].Cells[2].Value = Math.Round(val2, 2);

                        double comp4;
                        double val4 = 100000;
                        foreach (DataGridViewRow max4 in dataGridView3.Rows)
                        {
                            string max41 = Convert.ToString(max4.Cells[5].Value);
                            comp4 = Convert.ToDouble(max41);
                            if (comp4 < val4)
                            {
                                val4 = comp4;
                                lname = max4.Cells[0].Value.ToString();
                                dataGridView4.Rows[n_filas - 1].Cells[0].Value = lname;
                                dataGridView6.Rows[n_filas - 1].Cells[0].Value = lname;
                            }
                        }
                        dataGridView4.Rows[n_filas - 1].Cells[2].Value = Math.Round(val4, 2);
                        dataGridView6.Rows[n_filas - 1].Cells[2].Value = Math.Round(val4, 2);
                        //segunda carrera max%
                        double comp2z;
                        double val2z = 100000;
                        foreach (DataGridViewRow max21 in dataGridView14.Rows)
                        {
                            string max2 = Convert.ToString(max21.Cells[3].Value);
                            comp2z = Convert.ToDouble(max2);
                            if (comp2z < val2z)
                            {
                                val2z = comp2z;
                            }
                        }
                        dataGridView15.Rows[0].Cells[2].Value = Math.Round(val2z, 2);
                        dataGridView12.Rows[0].Cells[2].Value = Math.Round(val2z, 2);

                        double comp4z;
                        double val4z = 100000;
                        foreach (DataGridViewRow max4 in dataGridView14.Rows)
                        {
                            string max41 = Convert.ToString(max4.Cells[5].Value);
                            comp4z = Convert.ToDouble(max41);
                            if (comp4z < val4z)
                            {
                                val4z = comp4z;
                                lnamez = max4.Cells[0].Value.ToString();
                                dataGridView15.Rows[n_filas1 - 1].Cells[0].Value = lnamez;
                                dataGridView12.Rows[n_filas1 - 1].Cells[0].Value = lnamez;
                            }
                        }
                        dataGridView15.Rows[n_filas1 - 1].Cells[2].Value = Math.Round(val4z, 2);
                        dataGridView12.Rows[n_filas1 - 1].Cells[2].Value = Math.Round(val4z, 2);

                        //Para 95%
                        double acumulador = 0;
                        double acumulador1 = 0;
                        int com = 0;
                        int com1 = 0;
                        double calculo;
                        double calculo1;
                        int check = 1;
                        int check1 = 0;
                        //Para max%
                        double acumuladorz = 0;
                        double acumulador1z = 0;
                        int comz = 0;
                        int com1z = 0;
                        double calculoz;
                        double calculo1z;
                        int checkz = 1;
                        int check1z = 0;
                        //Llenar los ultimos datos 95%
                        foreach (DataGridViewRow con in dataGridView3.Rows)
                        {
                            //Primera Corrida
                            while (com < 1)
                            {
                                acumulador = Convert.ToDouble(dataGridView4.Rows[0].Cells[1].Value);
                                com++;
                            }
                            if (con.Cells[0].Value.ToString() != name.ToString())
                            {
                                //Operaciones
                                calculo = ((Convert.ToDouble(con.Cells[2].Value)) - (Convert.ToDouble(dataGridView3.Rows[check1].Cells[2].Value.ToString())));
                                dataGridView4.Rows[check].Cells[1].Value = Math.Round(calculo, 2);
                                dataGridView6.Rows[check].Cells[1].Value = Math.Round(calculo, 2);

                                //Aumento de acumulador
                                acumulador = acumulador + Convert.ToDouble(con.Cells[2].Value.ToString());
                            }

                            //Segunda Corrida
                            while (com1 < 1)
                            {
                                acumulador1 = Convert.ToDouble(dataGridView4.Rows[0].Cells[2].Value);
                                com1++;
                            }
                            if (con.Cells[0].Value.ToString() != name.ToString())
                            {
                                //Operaciones
                                calculo1 = ((Convert.ToDouble(con.Cells[3].Value)) - (Convert.ToDouble(dataGridView3.Rows[check1].Cells[3].Value.ToString())));
                                dataGridView4.Rows[check].Cells[2].Value = Math.Round(calculo1, 2);
                                dataGridView6.Rows[check].Cells[2].Value = Math.Round(calculo1, 2);

                                //Aumento de acumulador
                                acumulador1 = acumulador1 + Convert.ToDouble(con.Cells[3].Value.ToString());
                                check++;
                                check1++;
                            }
                        }

                        //Llenar los ultimos datos max%
                        foreach (DataGridViewRow con in dataGridView14.Rows)
                        {
                            //Primera Corrida
                            while (comz < 1)
                            {
                                acumuladorz = Convert.ToDouble(dataGridView15.Rows[0].Cells[1].Value);
                                comz++;
                            }
                            if (con.Cells[0].Value.ToString() != namez.ToString())
                            {
                                //Operaciones
                                calculoz = ((Convert.ToDouble(con.Cells[2].Value)) - (Convert.ToDouble(dataGridView14.Rows[check1z].Cells[2].Value.ToString())));
                                dataGridView15.Rows[checkz].Cells[1].Value = Math.Round(calculoz, 2);
                                dataGridView12.Rows[checkz].Cells[1].Value = Math.Round(calculoz, 2);

                                //Aumento de acumulador
                                acumuladorz = acumuladorz + Convert.ToDouble(con.Cells[2].Value.ToString());
                            }

                            //Segunda Corrida
                            while (com1z < 1)
                            {
                                acumulador1z = Convert.ToDouble(dataGridView15.Rows[0].Cells[2].Value);
                                com1z++;
                            }
                            if (con.Cells[0].Value.ToString() != namez.ToString())
                            {
                                //Operaciones
                                calculo1z = ((Convert.ToDouble(con.Cells[3].Value)) - (Convert.ToDouble(dataGridView14.Rows[check1z].Cells[3].Value.ToString())));
                                dataGridView15.Rows[checkz].Cells[2].Value = Math.Round(calculo1z, 2);
                                dataGridView12.Rows[checkz].Cells[2].Value = Math.Round(calculo1z, 2);

                                //Aumento de acumulador
                                acumulador1z = acumulador1z + Convert.ToDouble(con.Cells[3].Value.ToString());
                                checkz++;
                                check1z++;
                            }
                        }
                        renombrar1();
                    }
                    dataGridView4.Columns[3].Visible = false;
                    dataGridView15.Columns[3].Visible = false;
                }
                else if (dataGridView1.Rows[0].Cells[2].Value.ToString() == "Run_1 (Vol%)")
                {
                    //1 corrida
                    if (dataGridView3.Rows.Count == 2)
                    {
                        dataGridView4.Visible = true;
                        dataGridView15.Visible = true;
                        //Funciones del diferencial 95%
                        dataGridView4.Rows.Add();
                        dataGridView4.Rows.Add();
                        dataGridView4.Rows.Add();

                        dataGridView6.Rows.Add();
                        dataGridView6.Rows.Add();
                        dataGridView6.Rows.Add();
                        //Funciones del diferencial max%
                        dataGridView15.Rows.Add();
                        dataGridView15.Rows.Add();
                        dataGridView15.Rows.Add();

                        dataGridView12.Rows.Add();
                        dataGridView12.Rows.Add();
                        dataGridView12.Rows.Add();
                        //Asignacion de variables de comparacion 95%
                        double comp1;
                        double val1 = 100000;
                        foreach (DataGridViewRow max1 in dataGridView3.Rows)
                        {
                            string max11 = Convert.ToString(max1.Cells[2].Value);
                            comp1 = Convert.ToDouble(max11);
                            if (comp1 < val1)
                            {
                                val1 = comp1;
                                name = max1.Cells[0].Value.ToString();
                                dataGridView4.Rows[0].Cells[0].Value = name;
                                dataGridView6.Rows[0].Cells[0].Value = name;
                            }
                        }
                        dataGridView4.Rows[0].Cells[1].Value = Math.Round(val1, 2);
                        dataGridView6.Rows[0].Cells[1].Value = Math.Round(val1, 2);
                        //Asignacion de variables de comparacion max%
                        double comp1z;
                        double val1z = 100000;
                        foreach (DataGridViewRow max1 in dataGridView14.Rows)
                        {
                            string max11 = Convert.ToString(max1.Cells[2].Value);
                            comp1z = Convert.ToDouble(max11);
                            if (comp1z < val1z)
                            {
                                val1z = comp1z;
                                namez = max1.Cells[0].Value.ToString();
                                dataGridView15.Rows[0].Cells[0].Value = namez;
                                dataGridView12.Rows[0].Cells[0].Value = namez;
                            }
                        }
                        dataGridView15.Rows[0].Cells[1].Value = Math.Round(val1z, 2);
                        dataGridView12.Rows[0].Cells[1].Value = Math.Round(val1z, 2);

                        //4 95%
                        double comp4;
                        double val4 = 1000000;
                        string name2;
                        foreach (DataGridViewRow max4 in dataGridView3.Rows)
                        {
                            string max41 = Convert.ToString(max4.Cells[3].Value);
                            comp4 = Convert.ToDouble(max41);
                            if (comp4 < val4)
                            {
                                val4 = comp4;
                                name2 = max4.Cells[0].Value.ToString();
                                dataGridView4.Rows[2].Cells[0].Value = name2;
                                dataGridView6.Rows[2].Cells[0].Value = name2;
                            }
                        }
                        dataGridView4.Rows[2].Cells[1].Value = Math.Round(val4, 2);
                        dataGridView6.Rows[2].Cells[1].Value = Math.Round(val4, 2);
                        //4 max%
                        double comp4z;
                        double val4z = 1000000;
                        string name2z;
                        foreach (DataGridViewRow max4 in dataGridView14.Rows)
                        {
                            string max41 = Convert.ToString(max4.Cells[3].Value);
                            comp4z = Convert.ToDouble(max41);
                            if (comp4z < val4z)
                            {
                                val4z = comp4z;
                                name2z = max4.Cells[0].Value.ToString();
                                dataGridView15.Rows[2].Cells[0].Value = name2z;
                                dataGridView12.Rows[2].Cells[0].Value = name2z;
                            }
                        }
                        dataGridView15.Rows[2].Cells[1].Value = Math.Round(val4z, 2);
                        dataGridView12.Rows[2].Cells[1].Value = Math.Round(val4z, 2);

                        //Crear el diferencial 95%
                        dataGridView4.Rows[1].Cells[1].Value =
                        Math.Round(Convert.ToDouble(100 - (Convert.ToDouble(dataGridView4.Rows[2].Cells[1].Value) +
                        Convert.ToDouble(dataGridView4.Rows[0].Cells[1].Value))), 2);

                        dataGridView6.Rows[1].Cells[1].Value =
                        Math.Round(Convert.ToDouble(100 - (Convert.ToDouble(dataGridView4.Rows[2].Cells[1].Value) +
                        Convert.ToDouble(dataGridView4.Rows[0].Cells[1].Value))), 2);
                        //Crear el diferencial max%
                        dataGridView15.Rows[1].Cells[1].Value =
                        Math.Round(Convert.ToDouble(100 - (Convert.ToDouble(dataGridView15.Rows[2].Cells[1].Value) +
                        Convert.ToDouble(dataGridView15.Rows[0].Cells[1].Value))), 2);

                        dataGridView12.Rows[1].Cells[1].Value =
                        Math.Round(Convert.ToDouble(100 - (Convert.ToDouble(dataGridView15.Rows[2].Cells[1].Value) +
                        Convert.ToDouble(dataGridView15.Rows[0].Cells[1].Value))), 2);

                        renombrar();
                    }
                    else if (dataGridView3.Rows.Count > 2)
                    {
                        dataGridView4.Visible = true;
                        dataGridView15.Visible = true;

                        //95%
                        int n_filas = Convert.ToInt32(dataGridView3.Rows.Count.ToString()) + 1;
                        int contador = 1;
                        while (contador < n_filas)
                        {
                            dataGridView4.Rows.Add();
                            dataGridView6.Rows.Add();
                            contador++;
                        }
                        //max%
                        int n_filas1 = Convert.ToInt32(dataGridView14.Rows.Count.ToString()) + 1;
                        int contador1 = 1;
                        while (contador1 < n_filas1)
                        {
                            dataGridView15.Rows.Add();
                            dataGridView12.Rows.Add();
                            contador1++;
                        }

                        //llenar la primera y la ultima columna 95%
                        double comp;
                        double val = 100000;
                        foreach (DataGridViewRow max in dataGridView3.Rows)
                        {
                            string max1 = Convert.ToString(max.Cells[2].Value);
                            comp = Convert.ToDouble(max1);
                            if (comp < val)
                            {
                                val = comp;
                                name = max.Cells[0].Value.ToString();
                                dataGridView4.Rows[0].Cells[0].Value = name;
                                dataGridView6.Rows[0].Cells[0].Value = name;
                            }
                        }
                        dataGridView4.Rows[0].Cells[1].Value = Math.Round(val, 2);
                        dataGridView6.Rows[0].Cells[1].Value = Math.Round(val, 2);
                        //llenar la primera y la ultima columna max%
                        double comp1;
                        double val1 = 100000;
                        foreach (DataGridViewRow max in dataGridView14.Rows)
                        {
                            string max1 = Convert.ToString(max.Cells[2].Value);
                            comp1 = Convert.ToDouble(max1);
                            if (comp1 < val1)
                            {
                                val1 = comp1;
                                namez = max.Cells[0].Value.ToString();
                                dataGridView15.Rows[0].Cells[0].Value = namez;
                                dataGridView12.Rows[0].Cells[0].Value = namez;
                            }
                        }
                        dataGridView15.Rows[0].Cells[1].Value = Math.Round(val1, 2);
                        dataGridView12.Rows[0].Cells[1].Value = Math.Round(val1, 2);

                        //2 95%
                        double comp1z;
                        double val1z = 100000;
                        foreach (DataGridViewRow max1 in dataGridView3.Rows)
                        {
                            string max11 = Convert.ToString(max1.Cells[3].Value);
                            comp1z = Convert.ToDouble(max11);
                            if (comp1z < val1z)
                            {
                                val1z = comp1z;
                            }
                        }
                        dataGridView4.Rows[n_filas - 1].Cells[1].Value = Math.Round(val1z, 2);
                        dataGridView6.Rows[n_filas - 1].Cells[1].Value = Math.Round(val1z, 2);
                        //2 max%
                        double comp1z1;
                        double val1z1 = 100000;
                        foreach (DataGridViewRow max1 in dataGridView14.Rows)
                        {
                            string max11 = Convert.ToString(max1.Cells[3].Value);
                            comp1z1 = Convert.ToDouble(max11);
                            if (comp1z1 < val1z1)
                            {
                                val1z1 = comp1z1;
                            }
                        }
                        dataGridView15.Rows[n_filas1 - 1].Cells[1].Value = Math.Round(val1z1, 2);
                        dataGridView12.Rows[n_filas1 - 1].Cells[1].Value = Math.Round(val1z1, 2);

                        //95%
                        double acumulador = 0;
                        int com = 0;
                        double calculo;
                        int check = 1;
                        int check1 = 0;
                        //95%
                        double acumuladorz = 0;
                        int comz = 0;
                        double calculoz;
                        int checkz = 1;
                        int check1z = 0;

                        //Llenar los ultimos datos 95%
                        foreach (DataGridViewRow con in dataGridView3.Rows)
                        {
                            //Primera Corrida
                            while (com < 1)
                            {
                                acumulador = Convert.ToDouble(dataGridView4.Rows[0].Cells[1].Value);
                                com++;
                            }
                            if (con.Cells[0].Value.ToString() != name.ToString())
                            {
                                //Operaciones
                                calculo = ((Convert.ToDouble(con.Cells[2].Value)) - (Convert.ToDouble(dataGridView3.Rows[check1].Cells[2].Value.ToString())));
                                dataGridView4.Rows[check].Cells[1].Value = Math.Round(calculo, 2);
                                dataGridView6.Rows[check].Cells[1].Value = Math.Round(calculo, 2);

                                //Aumento de acumulador
                                acumulador = acumulador + Convert.ToDouble(con.Cells[2].Value.ToString());
                                check++;
                                check1++;
                            }
                        }
                        //Llenar los ultimos datos max%
                        foreach (DataGridViewRow con in dataGridView14.Rows)
                        {
                            //Primera Corrida
                            while (comz < 1)
                            {
                                acumuladorz = Convert.ToDouble(dataGridView15.Rows[0].Cells[1].Value);
                                comz++;
                            }
                            if (con.Cells[0].Value.ToString() != namez.ToString())
                            {
                                //Operaciones
                                calculoz = ((Convert.ToDouble(con.Cells[2].Value)) - (Convert.ToDouble(dataGridView14.Rows[check1z].Cells[2].Value.ToString())));
                                dataGridView15.Rows[checkz].Cells[1].Value = Math.Round(calculoz, 2);
                                dataGridView12.Rows[checkz].Cells[1].Value = Math.Round(calculoz, 2);

                                //Aumento de acumulador
                                acumuladorz = acumuladorz + Convert.ToDouble(con.Cells[2].Value.ToString());
                                checkz++;
                                check1z++;
                            }
                        }
                        renombrar1();
                    }
                    dataGridView4.Columns[3].Visible = false;
                    dataGridView4.Columns[2].Visible = false;
                    dataGridView15.Columns[3].Visible = false;
                    dataGridView15.Columns[2].Visible = false;
                }
            }
            catch (Exception tr)
            {
                dataGridView4.Visible = false;
                dataGridView15.Visible = false;
                MessageBox.Show("It's necessary to mark the cumulative");
            }

            dataGridView5.Visible = true;
            dataGridView6.Visible = true;
            button7.Visible = true;
            allowSelect = true;
            tabControl1.SelectedTab = tabPage3;
            allowSelect = false;
            dataGridView5.AllowUserToAddRows = false;

            dataGridView11.Visible = true;
            dataGridView12.Visible = true;
            dataGridView11.AllowUserToAddRows = false;

            //Asignacion a variables
            string n1;
            string n20;
            string n30;

            string f1;
            string f2;
            string f3;

            string u1;
            string u2;
            string u3;

            string e1;
            string e2;
            string e3;

            string i1;
            string i2;
            string i3;

            string g1;
            string g2;
            string g3;

            string l1;
            string l2;
            string l3;

            string c1;
            string c2;
            string c3;

            string cl1;
            string cl2;
            string cl3;

            //Busqueda de datos de la empresa
            foreach (DataGridViewRow row in dataGridView7.Rows)
            {
                try
                {
                    if (row.Cells[1].Value.ToString() == "Name")
                    {
                        Nombres.Add(row.Cells[2].Value.ToString());
                    }
                    if (row.Cells[1].Value.ToString() == "Sample Date")
                    {
                        Fecha.Add(row.Cells[2].Value.ToString());
                    }
                    if (row.Cells[1].Value.ToString() == "User")
                    {
                        Usuarios.Add(row.Cells[2].Value.ToString());
                    }
                    if (row.Cells[1].Value.ToString() == "Device")
                    {
                        Equipos.Add(row.Cells[2].Value.ToString());
                    }
                    if (row.Cells[1].Value.ToString() == "Sample ID")
                    {
                        Ids.Add(row.Cells[2].Value.ToString());
                    }
                    if (row.Cells[1].Value.ToString() == "Group ID")
                    {
                        Grupos.Add(row.Cells[2].Value.ToString());
                    }
                    if (row.Cells[1].Value.ToString() == "Batch")
                    {
                        Lotes.Add(row.Cells[2].Value.ToString());
                    }
                    else
                    {
                        Lotes.Add("");
                    }
                    if (row.Cells[1].Value.ToString() == "Comments")
                    {
                        Comentarios.Add(row.Cells[2].Value.ToString());
                    }
                    else
                    {
                        Comentarios.Add("");
                    }
                    if (row.Cells[1].Value.ToString() == "Customer")
                    {
                        Clientes.Add(row.Cells[2].Value.ToString());
                    }
                    else
                    {
                        Clientes.Add("");
                    }
                }
                catch (Exception sw)
                {

                }
            }

            //Busqueda de si son 3 o menos corridas
            if (dataGridView1.Rows[0].Cells[4].Value.ToString() == "Run_3 (Vol%)")
            {
                //3 corridas
                //Asignacion a variables
                if (Nombres.Count > 0)
                {
                    n1 = Nombres[0];
                    n20 = Nombres[1];
                    n30 = Nombres[2];

                    label14.Text = n1;
                    label31.Text = n20;
                    label49.Text = n30;
                }
                else
                {
                    label14.Text = "";
                    label31.Text = "";
                    label49.Text = "";

                    label4.Visible = false;
                    label40.Visible = false;
                    label58.Visible = false;
                }
                if (Fecha.Count > 0)
                {
                    f1 = Fecha[0];
                    f2 = Fecha[1];
                    f3 = Fecha[2];

                    label15.Text = f1;
                    label30.Text = f2;
                    label48.Text = f3;
                }
                else
                {
                    label15.Text = "";
                    label30.Text = "";
                    label48.Text = "";
                }
                if (Usuarios.Count > 0)
                {
                    u1 = Usuarios[0];
                    u2 = Usuarios[1];
                    u3 = Usuarios[2];

                    label16.Text = u1;
                    label29.Text = u2;
                    label47.Text = u3;
                }
                else
                {
                    label16.Text = "";
                    label29.Text = "";
                    label47.Text = "";
                }
                if (Equipos.Count > 0)
                {
                    e1 = Equipos[0];
                    e2 = Equipos[1];
                    e3 = Equipos[2];

                    label17.Text = e1;
                    label28.Text = e2;
                    label46.Text = e3;
                }
                else
                {
                    label17.Text = "";
                    label28.Text = "";
                    label46.Text = "";
                }
                if (Ids.Count > 0)
                {
                    i1 = Ids[0];
                    i2 = Ids[1];
                    i3 = Ids[2];

                    label18.Text = i1;
                    label27.Text = i2;
                    label45.Text = i3;
                }
                else
                {
                    label18.Text = "";
                    label27.Text = "";
                    label45.Text = "";
                }
                if (Grupos.Count > 0)
                {
                    g1 = Grupos[0];
                    g2 = Grupos[1];
                    g3 = Grupos[2];

                    label19.Text = g1;
                    label26.Text = g2;
                    label44.Text = g3;
                }
                else
                {
                    label19.Text = "";
                    label26.Text = "";
                    label44.Text = "";
                }
                if (Lotes.Count > 0)
                {
                    l1 = Lotes[0];
                    l2 = Lotes[1];
                    l3 = Lotes[2];

                    label20.Text = l1;
                    label25.Text = l2;
                    label43.Text = l3;

                    label10.Text = "Batch: ";
                    label34.Text = "Batch: ";
                    label52.Text = "Batch: ";
                    if (l1 == "")
                    {
                        label10.Text = "";
                    }
                    if (l2 == "")
                    {
                        label34.Text = "";
                    }
                    if (l3 == "")
                    {
                        label52.Text = "";
                    }
                }
                else
                {
                    label20.Text = "";
                    label25.Text = "";
                    label43.Text = "";

                    label10.Text = "";
                    label34.Text = "";
                    label52.Text = "";
                }
                if (Comentarios.Count > 0)
                {
                    c1 = Comentarios[0];
                    c2 = Comentarios[1];
                    c3 = Comentarios[2];

                    label21.Text = c1;
                    label24.Text = c2;
                    label42.Text = c3;

                    label11.Text = "Comments: ";
                    label33.Text = "Comments: ";
                    label51.Text = "Comments: ";
                    if (c1 == "")
                    {
                        label11.Text = "";
                    }
                    if (c2 == "")
                    {
                        label33.Text = "";
                    }
                    if (c3 == "")
                    {
                        label51.Text = "";
                    }
                }
                else
                {
                    label21.Text = "";
                    label24.Text = "";
                    label42.Text = "";

                    label11.Text = "";
                    label33.Text = "";
                    label51.Text = "";
                }
                if (Clientes.Count > 0)
                {
                    cl1 = Clientes[0];
                    cl2 = Clientes[1];
                    cl3 = Clientes[2];

                    label22.Text = cl1;
                    label23.Text = cl2;
                    label41.Text = cl3;

                    label12.Text = "Customers: ";
                    label32.Text = "Customers: ";
                    label50.Text = "Customers: ";
                    if (cl1 == "")
                    {
                        label12.Text = "";
                    }
                    if (cl2 == "")
                    {
                        label32.Text = "";
                    }
                    if (cl3 == "")
                    {
                        label50.Text = "";
                    }
                }
                else
                {
                    label22.Text = "";
                    label23.Text = "";
                    label41.Text = "";

                    label12.Text = "";
                    label32.Text = "";
                    label50.Text = "";
                }
            }
            else if (dataGridView1.Rows[0].Cells[3].Value.ToString() == "Run_2 (Vol%)")
            {
                // 2 corridas
                //Asignacion a variables
                if (Nombres.Count > 0)
                {
                    n1 = Nombres[0];
                    n20 = Nombres[1];

                    label14.Text = n1;
                    label31.Text = n20;
                    label49.Text = "";
                    label60.Text = "";
                    label58.Text = "";
                }
                else
                {
                    label14.Text = "";
                    label31.Text = "";
                    label49.Text = "";

                    label4.Visible = false;
                    label40.Visible = false;
                    label58.Visible = false;
                }
                if (Fecha.Count > 0)
                {
                    f1 = Fecha[0];
                    f2 = Fecha[1];

                    label15.Text = f1;
                    label30.Text = f2;
                    label48.Text = "";
                    label57.Text = "";
                }
                else
                {
                    label15.Text = "";
                    label30.Text = "";
                    label48.Text = "";
                }
                if (Usuarios.Count > 0)
                {
                    u1 = Usuarios[0];
                    u2 = Usuarios[1];

                    label16.Text = u1;
                    label29.Text = u2;
                    label47.Text = "";
                    label56.Text = "";
                }
                else
                {
                    label16.Text = "";
                    label29.Text = "";
                    label47.Text = "";
                }
                if (Equipos.Count > 0)
                {
                    e1 = Equipos[0];
                    e2 = Equipos[1];

                    label17.Text = e1;
                    label28.Text = e2;
                    label46.Text = "";
                    label55.Text = "";
                }
                else
                {
                    label17.Text = "";
                    label28.Text = "";
                    label46.Text = "";
                }
                if (Ids.Count > 0)
                {
                    i1 = Ids[0];
                    i2 = Ids[1];

                    label18.Text = i1;
                    label27.Text = i2;
                    label45.Text = "";
                    label54.Text = "";
                }
                else
                {
                    label18.Text = "";
                    label27.Text = "";
                    label45.Text = "";
                }
                if (Grupos.Count > 0)
                {
                    g1 = Grupos[0];
                    g2 = Grupos[1];

                    label19.Text = g1;
                    label26.Text = g2;
                    label44.Text = "";
                    label53.Text = "";
                }
                else
                {
                    label19.Text = "";
                    label26.Text = "";
                    label44.Text = "";
                }
                if (Lotes.Count > 0)
                {
                    l1 = Lotes[0];
                    l2 = Lotes[1];

                    label20.Text = l1;
                    label25.Text = l2;
                    label43.Text = "";

                    label10.Text = "Batch: ";
                    label34.Text = "Batch: ";
                    label52.Text = "";
                    if (l1 == "")
                    {
                        label10.Text = "";
                    }
                    if (l2 == "")
                    {
                        label34.Text = "";
                    }
                }
                else
                {
                    label20.Text = "";
                    label25.Text = "";
                    label43.Text = "";

                    label10.Text = "";
                    label34.Text = "";
                    label52.Text = "";
                }
                if (Comentarios.Count > 0)
                {
                    c1 = Comentarios[0];
                    c2 = Comentarios[1];

                    label21.Text = c1;
                    label24.Text = c2;
                    label42.Text = "";

                    label11.Text = "Comments: ";
                    label33.Text = "Comments: ";
                    label51.Text = "";
                    if (c1 == "")
                    {
                        label11.Text = "";
                    }
                    if (c2 == "")
                    {
                        label33.Text = "";
                    }
                }
                else
                {
                    label21.Text = "";
                    label24.Text = "";
                    label42.Text = "";

                    label11.Text = "";
                    label33.Text = "";
                    label51.Text = "";
                }
                if (Clientes.Count > 0)
                {
                    cl1 = Clientes[0];
                    cl2 = Clientes[1];

                    label22.Text = cl1;
                    label23.Text = cl2;
                    label41.Text = "";

                    label12.Text = "Customers: ";
                    label32.Text = "Customers: ";
                    label50.Text = "";
                    label87.Text = "";
                    if (cl1 == "")
                    {
                        label12.Text = "";
                    }
                    if (cl2 == "")
                    {
                        label32.Text = "";
                    }
                }
                else
                {
                    label22.Text = "";
                    label23.Text = "";
                    label41.Text = "";

                    label12.Text = "";
                    label32.Text = "";
                    label50.Text = "";
                }
            }
            else if (dataGridView1.Rows[0].Cells[2].Value.ToString() == "Run_1 (Vol%)")
            {
                //1 corrida
                //Asignacion a variables
                if (Nombres.Count > 0)
                {
                    n1 = Nombres[0];

                    label14.Text = n1;
                    label31.Text = "";
                    label49.Text = "";
                    label40.Text = "";
                    label58.Text = "";
                }
                else
                {
                    label14.Text = "";
                    label31.Text = "";
                    label49.Text = "";

                    label4.Visible = false;
                    label40.Visible = false;
                    label58.Visible = false;
                }
                if (Fecha.Count > 0)
                {
                    f1 = Fecha[0];

                    label15.Text = f1;
                    label30.Text = "";
                    label48.Text = "";
                    label39.Text = "";
                    label57.Text = "";
                }
                else
                {
                    label15.Text = "";
                    label30.Text = "";
                    label48.Text = "";
                }
                if (Usuarios.Count > 0)
                {
                    u1 = Usuarios[0];

                    label16.Text = u1;
                    label29.Text = "";
                    label47.Text = "";
                    label38.Text = "";
                    label56.Text = "";
                }
                else
                {
                    label16.Text = "";
                    label29.Text = "";
                    label47.Text = "";
                }
                if (Equipos.Count > 0)
                {
                    e1 = Equipos[0];

                    label17.Text = e1;
                    label28.Text = "";
                    label46.Text = "";
                    label37.Text = "";
                    label55.Text = "";
                }
                else
                {
                    label17.Text = "";
                    label28.Text = "";
                    label46.Text = "";
                }
                if (Ids.Count > 0)
                {
                    i1 = Ids[0];

                    label18.Text = i1;
                    label27.Text = "";
                    label45.Text = "";
                    label36.Text = "";
                    label54.Text = "";
                }
                else
                {
                    label18.Text = "";
                    label27.Text = "";
                    label45.Text = "";
                }
                if (Grupos.Count > 0)
                {
                    g1 = Grupos[0];

                    label19.Text = g1;
                    label26.Text = "";
                    label44.Text = "";
                    label35.Text = "";
                    label53.Text = "";
                }
                else
                {
                    label19.Text = "";
                    label26.Text = "";
                    label44.Text = "";
                }
                if (Lotes.Count > 0)
                {
                    l1 = Lotes[0];

                    label20.Text = l1;
                    label25.Text = "";
                    label43.Text = "";

                    label10.Text = "Batch: ";
                    label34.Text = "";
                    label52.Text = "";
                    if (l1 == "")
                    {
                        label10.Text = "";
                    }
                }
                else
                {
                    label20.Text = "";
                    label25.Text = "";
                    label43.Text = "";

                    label10.Text = "";
                    label34.Text = "";
                    label52.Text = "";
                }
                if (Comentarios.Count > 0)
                {
                    c1 = Comentarios[0];

                    label21.Text = c1;
                    label24.Text = "";
                    label42.Text = "";

                    label11.Text = "Comments: ";
                    label33.Text = "";
                    label51.Text = "";
                    if (c1 == "")
                    {
                        label11.Text = "";
                    }
                }
                else
                {
                    label21.Text = "";
                    label24.Text = "";
                    label42.Text = "";

                    label11.Text = "";
                    label33.Text = "";
                    label51.Text = "";
                }
                if (Clientes.Count > 0)
                {
                    cl1 = Clientes[0];

                    label22.Text = cl1;
                    label23.Text = "";
                    label41.Text = "";

                    label12.Text = "Customers: ";
                    label32.Text = "";
                    label50.Text = "";
                    if (cl1 == "")
                    {
                        label12.Text = "";
                    }
                }
                else
                {
                    label22.Text = "";
                    label23.Text = "";
                    label41.Text = "";

                    label12.Text = "";
                    label32.Text = "";
                    label50.Text = "";
                }
                label59.Text = "";
                label60.Text = "";
            }


            //Asignacion de Datos de la empresa 95%
            dataGridView5.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridView5.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;

            dataGridView6.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridView6.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;

            //Asignacion de Datos de la empresa max%
            dataGridView11.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridView11.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;

            dataGridView12.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridView12.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;

            //Insertar 999 en dgv 5
            dataGridView5.Rows.Insert(0, "", "999", "0", "0", "0", "0", "0", "0");
            dataGridView11.Rows.Insert(0, "", "999", "0", "0", "0", "0", "0", "0");

            dataGridView5.ReadOnly = true;
            dataGridView6.ReadOnly = true;
            dataGridView5.ClearSelection();
            dataGridView6.ClearSelection();

            dataGridView11.ReadOnly = true;
            dataGridView12.ReadOnly = true;
            dataGridView11.ClearSelection();
            dataGridView12.ClearSelection();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            //Oculta el diferencial del reporte
            con_dif = "si";
            try
            {
                foreach (DataGridViewRow row in dataGridView6.Rows)
                {
                    foreach (DataGridViewColumn col in dataGridView6.Columns)
                    {
                        dataGridView6.Rows[row.Index].Cells[col.Index].Value = "";
                    }
                }

                foreach (DataGridViewRow row in dataGridView12.Rows)
                {
                    foreach (DataGridViewColumn col in dataGridView12.Columns)
                    {
                        dataGridView12.Rows[row.Index].Cells[col.Index].Value = "";
                    }
                }
            }
            catch (Exception df)
            {
            }
            button8.Visible = false;
            dataGridView6.Visible = false;
            dataGridView12.Visible = false;
        }
    }
}
