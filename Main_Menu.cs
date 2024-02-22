using System;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using System.IO;
using ExcelDataReader;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using MaterialSkin.Controls;
using System.Diagnostics;
using LecturaExcel.Responsabilitis;
using LecturaExcel.View;
using LecturaExcel.Controller;

namespace LecturaExcel
{
    public partial class Main_Menu : MaterialForm
    {
        public Main_Menu(string filePath)
        {
            InitializeComponent();
            C_LoadFile controller = new C_LoadFile();
            controller.readExcelFile(filePath);

            Dgv_Particle_Data.DataSource = controller.getParticleData();
            Dgv_Sample_Information.DataSource = controller.getSampleInformation();
        }

        //Declaracion de variables 
        string detectorNumberASTM95;
        string detectorNumberSingleAperture;
        string name;
        string namez;
        string lname;
        string lnamez;
        int check = 0;
        int corridas = 3;

        bool ch1 = true;
        bool ch2 = true;
        bool allowSelect = false;
        bool oc1 = true;
        string numberOfRun = "";
        string accumulated = "no";
        string diffferential = "no";

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
        List<string> valor_nominal= new List<string>();

        private void Form1_Load(object sender, EventArgs e)
        {
            //Maximizar el tamaño de la ventana del form
            WindowState = FormWindowState.Maximized;
        }
        private void Btn_Load_File_Click(object sender, EventArgs e)
        {
            C_LoadFile controller = new C_LoadFile();
            controller.controll();
            Dgv_Particle_Data.DataSource = controller.getParticleData();
            Dgv_Sample_Information.DataSource = controller.getSampleInformation();
        }

#region PageMesh_Selection

        private void Btn_Go_To_Manual_Mesh_Selection_Click(object sender, EventArgs e)
        {
            valor_nominal.Add("-");
            Manage_Data data = new Manage_Data();
            try
            {
                allowSelect = true;
                TabControl_Main_Menu.SelectedTab = Page_Mesh_Selection;
                allowSelect = false;

                data.createDgvToleranceTableReference(Dgv_Tolerance_Table_Reference);
                data.addColumnToDatagridView("Particle Size", Dgv_ASTM95_Detector_Number);
                data.addColumnToDatagridView("Mesh #", Dgv_ASTM95_Detector_Number);
                data.addColumnToDatagridView("Values To Calculate", Dgv_ASTM95_Detector_Number);
                data.addColumnToDatagridView("Detector Number", Dgv_ASTM95_Detector_Number);
                data.addColumnToDatagridView("Particle Size", Dgv_Single_Aperture_Detector);
                data.addColumnToDatagridView("Mesh #", Dgv_Single_Aperture_Detector);
                data.addColumnToDatagridView("Values To Calculate", Dgv_Single_Aperture_Detector);
                data.addColumnToDatagridView("Record", Dgv_Single_Aperture_Detector);
                data.copyStructureOfDataGridViewToOther(Dgv_Particle_Data, Dgv_MAX_D95_Selected_Row);
                data.copyStructureOfDataGridViewToOther(Dgv_Particle_Data, Dgv_Single_Aperture_Selected_Row);
            }
            catch (Exception ex)
            {
                MessageBox.Show(" Please select an Excel file to continue "+ex.Message.ToString());
            }

            this.cleanOldInformationOfDataGridViews();

            ch1 = true;
            ch2 = true;
            data.removeUselessGridColumns(Dgv_ASTM_D95_Accumulated_rigth_left);
            data.removeUselessGridColumns(Dgv_ASTM_D95);
            data.removeUselessGridColumns(Dgv_Single_Aperture_Accumulated_right_left);
            data.removeUselessGridColumns(Dgv_ASTM_Single_Aperture);        
        }

        /*
        private void addUserControll()
        {
            panelTest.Controls.Clear();
            panelTest.Controls.Add(new UserControl1());
        }*/

        private void cleanOldInformationOfDataGridViews()
        {
            Dgv_ASTM95_Detector_Number.ReadOnly = true;
            Dgv_Single_Aperture_Detector.ReadOnly = true;
            Dgv_ASTM95_Detector_Number.Rows.Clear();
            Dgv_ASTM_D95.Rows.Clear();
            Dgv_ASTM95_Run_Differentials.Rows.Clear();
            Dgv_ASTM_D95_Accumulated_rigth_left.Rows.Clear();
            Dgv_Single_Aperture_Accumulated_right_left.Rows.Clear();
            Dgv_ASTM_95_Differential.Rows.Clear();
            Dgv_Single_Aperture_Differential.Rows.Clear();
            Dgv_ASTM_Single_Aperture.Rows.Clear();
            Dgv_Single_Aperture_Run_Differential.Rows.Clear();
        }
        
        private void ComboBox_Mesh_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Despues de seleccionar una malla que se quiera conocer sus datos dentro del excel, se hace una lista de referencia correspondiendo a lo que hay en el grid de el excel que se subio y con los datos de referencia que marcan los limites de cada malla
            string micronsSingleAperture = Dgv_Tolerance_Table_Reference.Rows[Combo_Box_Mesh.SelectedIndex].Cells[2].Value.ToString();
            string selectedMesh = Combo_Box_Mesh.SelectedItem.ToString();
            string micronsMax95 = Dgv_Tolerance_Table_Reference.Rows[Combo_Box_Mesh.SelectedIndex].Cells[3].Value.ToString();

            Dgv_ASTM95_Detector_Number.Rows.Add(micronsSingleAperture,
                selectedMesh, micronsSingleAperture, true);

            Dgv_Single_Aperture_Detector.Rows.Add(micronsMax95,
                selectedMesh, micronsMax95, true);

            Dgv_ASTM_D95.Rows.Add(micronsSingleAperture,
                selectedMesh, micronsSingleAperture, true);

            Dgv_ASTM_Single_Aperture.Rows.Add(micronsMax95,
                selectedMesh, micronsMax95, true);

            Dgv_ASTM_D95_Accumulated_rigth_left.Rows.Add(micronsSingleAperture,    
                selectedMesh, micronsSingleAperture, true);

            Dgv_Single_Aperture_Accumulated_right_left.Rows.Add(micronsMax95,
                selectedMesh, micronsMax95, true);

            Dgv_MAX_D95_Selected_Row.Rows.Clear();
            Manage_Data data = new Manage_Data();
            data.copyStructureOfDataGridViewToOther(Dgv_Particle_Data, Dgv_MAX_D95_Selected_Row);
            
            Dgv_Single_Aperture_Selected_Row.Rows.Clear();

            data.copyStructureOfDataGridViewToOther(Dgv_Particle_Data, Dgv_Single_Aperture_Selected_Row);
            data.copyStructureOfDataGridViewToOther(Dgv_Particle_Data, Dgv_MAX_D95_Selected_Row);
            Micron _micron = new Micron();                
            try
            {              
                //Ya que hace la busqueda del valor mas cercano al de la lista de referencia lo coloca en el Grid2
                foreach (DataGridViewRow Row in Dgv_ASTM95_Detector_Number.Rows)
                {
                    check = 0;
                    string micronInString = Convert.ToString(Row.Cells[2].Value);
                    double micron = _micron.getRoundedMicron(micronInString);
                    this.serchMoreNearValueInReferenceList(micron, Dgv_MAX_D95_Selected_Row, "detectorNumberASTM95");
                    Row.Cells[3].Value = detectorNumberASTM95;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }         
            try
            {
                //Ya que hace la busqueda del valor mas cercano al de la lista de referencia lo coloca en el Grid13
                foreach (DataGridViewRow Row in Dgv_Single_Aperture_Detector.Rows) 
                {
                    check = 0;
                    string micronInString = Convert.ToString(Row.Cells[2].Value);
                    double micron = _micron.getRoundedMicron(micronInString);
                    this.serchMoreNearValueInReferenceList(micron, Dgv_Single_Aperture_Selected_Row, "detectorNumberSingleAperture");
                    Row.Cells[3].Value = detectorNumberSingleAperture;
                }                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            valor_nominal.Add(Dgv_Tolerance_Table_Reference.Rows[Combo_Box_Mesh.SelectedIndex].Cells[0].Value.ToString());
        }

        private void serchMoreNearValueInReferenceList( double micron, DataGridView dataGridViewToFill, string rowToFill)
        {
            Micron _micron = new Micron();
            bool micronHasIntegers = _micron.checkIfMicronHasIntegers(micron);
            if (micronHasIntegers)
            {
                double roundedMicrons = micron * 1000;
                roundedMicrons = Math.Round(roundedMicrons, 0);
                this.serchForMicronValueInLowerLimit(roundedMicrons, dataGridViewToFill, rowToFill);
            }
            else
            {
                this.serchForMicronValueInLowerLimit(micron, dataGridViewToFill, rowToFill);
            }
        }

        private void serchForMicronValueInLowerLimit(double micron, DataGridView dataGridViewToFill, string rowToFill)
        {
            this.bringTheDataOfTheValueClosestToTheMicron(micron.ToString(), dataGridViewToFill, rowToFill);
            double lowerLimit = micron;
            while (check == 0)
            {
                lowerLimit = lowerLimit - 1;
                this.bringTheDataOfTheValueClosestToTheMicron(lowerLimit.ToString(), dataGridViewToFill, rowToFill);
            }
        }
        
        public void bringTheDataOfTheValueClosestToTheMicron(string micron, DataGridView dataGridViewToFill, string rowToFill)
        {
            string row;
            string valueOfCell;
            foreach (DataGridViewRow Row in Dgv_Particle_Data.Rows)
            {
                row = Row.Index.ToString();
                valueOfCell = Convert.ToString(Row.Cells[0].Value);

                if (valueOfCell != "Particle Size (µm)" && valueOfCell != "")
                {
                    double value = Convert.ToDouble(valueOfCell);
                    value = value + .4;
                    double roundedValue = Math.Round(value, 0);                
                    if (roundedValue.ToString() == micron)
                    {
                        dataGridViewToFill.Rows.Add(Dgv_Particle_Data.Rows[Convert.ToInt32(row)].Cells[0].Value.ToString(), Dgv_Particle_Data.Rows[Convert.ToInt32(row)].Cells[2].Value.ToString(), 
                            Dgv_Particle_Data.Rows[Convert.ToInt32(row)].Cells[3].Value.ToString(), Dgv_Particle_Data.Rows[Convert.ToInt32(row)].Cells[4].Value.ToString());

                        if (rowToFill.Equals("detectorNumberSingleAperture"))
                        {
                            detectorNumberSingleAperture = row;
                        }
                        if (rowToFill.Equals("detectorNumberASTM95"))
                        {
                            detectorNumberASTM95 = row;
                        }
                        check = 1;
                    }
                }
            }
        }

        #endregion
        private void Return_To_Mesh_Selection_Click(object sender, EventArgs e)
        {
            allowSelect = true;
            TabControl_Main_Menu.SelectedTab = Page_Mesh_Selection;
            allowSelect = false;

            Dgv_ASTM_D95_Accumulated_rigth_left.Rows.Clear();
            Dgv_ASTM_95_Differential.Rows.Clear();
        }

        private void Hide_Cumulatives_Click(object sender, EventArgs e)
        {
            accumulated = "si";
            Manage_Data data = new Manage_Data();
            data.hideCummulativeValues(numberOfRun, Dgv_ASTM_D95_Accumulated_rigth_left);
            data.hideCummulativeValues(numberOfRun, Dgv_Single_Aperture_Accumulated_right_left);
            Btn_Hide_Cumulatives.Visible = false;
        }

        private void tabControl1_Selecting(object sender, TabControlCancelEventArgs e)
        {
            //Para que el usuario no pueda cambiar el tabcontrol a placer y solo sea por medio de los botones
            if(!allowSelect) e.Cancel = true;
        }

        private void Btn_Generate_Report_Click(object sender, EventArgs e)
        {
            //Genera el reporte mandando a llamar a la funcion de finalizar
            this.generateReport();
        }

        //Aqui busco el que manda datos al reporte para modificarlo
        public void generateReport()
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

            while (Dgv_ASTM_D95_Accumulated_rigth_left.Rows.Count != valor_nominal.Count)
            {
                valor_nominal.Add("");
            }
            //If condicion de cuantas corridas son
            if (numberOfRun == "3")
            {
                //dgv5 = 8 columnas
                //Lectura de todos los datos para generar el reporte
                for (int i = 0; i < (Dgv_ASTM_D95_Accumulated_rigth_left.Rows.Count); i++)
                {
                    dt.Rows.Add(
                        Dgv_ASTM_D95_Accumulated_rigth_left.Rows[i].Cells[0].Value,
                        Dgv_ASTM_D95_Accumulated_rigth_left.Rows[i].Cells[1].Value,
                        Dgv_ASTM_D95_Accumulated_rigth_left.Rows[i].Cells[2].Value,
                        Dgv_ASTM_D95_Accumulated_rigth_left.Rows[i].Cells[3].Value,
                        Dgv_ASTM_D95_Accumulated_rigth_left.Rows[i].Cells[4].Value,
                        Dgv_ASTM_D95_Accumulated_rigth_left.Rows[i].Cells[5].Value,
                        Dgv_ASTM_D95_Accumulated_rigth_left.Rows[i].Cells[6].Value,
                        Dgv_ASTM_D95_Accumulated_rigth_left.Rows[i].Cells[7].Value,

                        Dgv_ASTM_95_Differential.Rows[i].Cells[1].Value,
                        Dgv_ASTM_95_Differential.Rows[i].Cells[2].Value,
                        Dgv_ASTM_95_Differential.Rows[i].Cells[3].Value,

                        (Lbl_Run_One.Text),
                        (Lbl_Name_Run_One.Text + Lbl_Name_Value_Run_One.Text),
                        (Lbl_Sample_Data_Run_One.Text + Lbl_Sample_Data_Value_Run_One.Text),
                        (Lbl_User_RunOne.Text + Lbl_User_Value_Run_One.Text),
                        (Lbl_Device_Run_One.Text + Lbl_Device_Value_Run_One.Text),
                        (Lbl_Sample_Id_Run_One.Text + Lbl_Sample_Id_Value_Run_One.Text),
                        (Lbl_Group_Id.Text + Lbl_Group_Id_Value.Text),
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

                        Dgv_Single_Aperture_Accumulated_right_left.Rows[i].Cells[2].Value,
                        Dgv_Single_Aperture_Accumulated_right_left.Rows[i].Cells[3].Value,
                        Dgv_Single_Aperture_Accumulated_right_left.Rows[i].Cells[4].Value,
                        Dgv_Single_Aperture_Accumulated_right_left.Rows[i].Cells[5].Value,
                        Dgv_Single_Aperture_Accumulated_right_left.Rows[i].Cells[6].Value,
                        Dgv_Single_Aperture_Accumulated_right_left.Rows[i].Cells[7].Value,

                        Dgv_Single_Aperture_Differential.Rows[i].Cells[1].Value,
                        Dgv_Single_Aperture_Differential.Rows[i].Cells[2].Value,
                        Dgv_Single_Aperture_Differential.Rows[i].Cells[3].Value,
                        Dgv_Single_Aperture_Accumulated_right_left.Rows[i].Cells[0].Value,
                        numberOfRun,
                        valor_nominal[i]
                        );
                }
                Vista_i vi = new Vista_i(dt);
                vi.Show();
            }
            else if (numberOfRun == "2")
            {
                //dgv5 = 6 columnas
                //Lectura de todos los datos para generar el reporte
                for (int i = 0; i < (Dgv_ASTM_D95_Accumulated_rigth_left.Rows.Count); i++)
                {
                    dt.Rows.Add(
                        Dgv_ASTM_D95_Accumulated_rigth_left.Rows[i].Cells[0].Value,
                        Dgv_ASTM_D95_Accumulated_rigth_left.Rows[i].Cells[1].Value,
                        Dgv_ASTM_D95_Accumulated_rigth_left.Rows[i].Cells[2].Value,
                        Dgv_ASTM_D95_Accumulated_rigth_left.Rows[i].Cells[3].Value,
                        Dgv_ASTM_D95_Accumulated_rigth_left.Rows[i].Cells[4].Value,
                        Dgv_ASTM_D95_Accumulated_rigth_left.Rows[i].Cells[5].Value,
                        "",
                        "",

                        Dgv_ASTM_95_Differential.Rows[i].Cells[1].Value,
                        Dgv_ASTM_95_Differential.Rows[i].Cells[2].Value,
                        "",

                        (Lbl_Run_One.Text),
                        (Lbl_Name_Run_One.Text + Lbl_Name_Value_Run_One.Text),
                        (Lbl_Sample_Data_Run_One.Text + Lbl_Sample_Data_Value_Run_One.Text),
                        (Lbl_User_RunOne.Text + Lbl_User_Value_Run_One.Text),
                        (Lbl_Device_Run_One.Text + Lbl_Device_Value_Run_One.Text),
                        (Lbl_Sample_Id_Run_One.Text + Lbl_Sample_Id_Value_Run_One.Text),
                        (Lbl_Group_Id.Text + Lbl_Group_Id_Value.Text),
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

                        Dgv_Single_Aperture_Accumulated_right_left.Rows[i].Cells[2].Value,
                        Dgv_Single_Aperture_Accumulated_right_left.Rows[i].Cells[3].Value,
                        Dgv_Single_Aperture_Accumulated_right_left.Rows[i].Cells[4].Value,
                        Dgv_Single_Aperture_Accumulated_right_left.Rows[i].Cells[5].Value,
                        "",
                        "",

                        Dgv_Single_Aperture_Differential.Rows[i].Cells[1].Value,
                        Dgv_Single_Aperture_Differential.Rows[i].Cells[2].Value,
                        "",
                        Dgv_Single_Aperture_Accumulated_right_left.Rows[i].Cells[0].Value,
                        numberOfRun,
                        valor_nominal[i]
                        );
                }
                Vista_i vi = new Vista_i(dt);
                vi.Show();
            }
            else if (numberOfRun == "1")
            {
                //dgv5 = 4 columnas
                //Lectura de todos los datos para generar el reporte
                for (int i = 0; i < (Dgv_ASTM_D95_Accumulated_rigth_left.Rows.Count); i++)
                {

                    dt.Rows.Add(
                        Dgv_ASTM_D95_Accumulated_rigth_left.Rows[i].Cells[0].Value,
                        Dgv_ASTM_D95_Accumulated_rigth_left.Rows[i].Cells[1].Value,
                        Dgv_ASTM_D95_Accumulated_rigth_left.Rows[i].Cells[2].Value,
                        Dgv_ASTM_D95_Accumulated_rigth_left.Rows[i].Cells[3].Value,
                        "",
                        "",
                        "",
                        "",

                        Dgv_ASTM_95_Differential.Rows[i].Cells[1].Value,
                        "",
                        "",

                        (Lbl_Run_One.Text),
                        (Lbl_Name_Run_One.Text + Lbl_Name_Value_Run_One.Text),
                        (Lbl_Sample_Data_Run_One.Text + Lbl_Sample_Data_Value_Run_One.Text),
                        (Lbl_User_RunOne.Text + Lbl_User_Value_Run_One.Text),
                        (Lbl_Device_Run_One.Text + Lbl_Device_Value_Run_One.Text),
                        (Lbl_Sample_Id_Run_One.Text + Lbl_Sample_Id_Value_Run_One.Text),
                        (Lbl_Group_Id.Text + Lbl_Group_Id_Value.Text),
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

                        Dgv_Single_Aperture_Accumulated_right_left.Rows[i].Cells[2].Value,
                        Dgv_Single_Aperture_Accumulated_right_left.Rows[i].Cells[3].Value,
                        "",
                        "",
                        "",
                        "",

                        Dgv_Single_Aperture_Differential.Rows[i].Cells[1].Value,
                        "",
                        "",
                        Dgv_Single_Aperture_Accumulated_right_left.Rows[i].Cells[0].Value,
                        numberOfRun,
                        valor_nominal[i]
                        );
                }
                Vista_i vi = new Vista_i(dt);
                vi.Show();
            }
        }

        private void Go_To_Report_View_Click(object sender, EventArgs e)
        {
            if (Dgv_Particle_Data.Rows[0].Cells[4].Value.ToString() == "Run_3 (Vol%)")
            {
                Manage_Data manageData = new Manage_Data();
                numberOfRun = "3";

                this.addColumnsOfCummulativeValues(Dgv_ASTM_D95, numberOfRun);
                this.addColumnsOfCummulativeValues(Dgv_ASTM_Single_Aperture, numberOfRun);
                this.addColumnsOfCummulativeValues(Dgv_ASTM_D95_Accumulated_rigth_left, numberOfRun);
                this.addColumnsOfCummulativeValues(Dgv_Single_Aperture_Accumulated_right_left, numberOfRun);

                //Aqui ira el calculo de la interpolacion de valores para 95%
                this.addCumulativeValuesForEachRunToDataGridView();

                ch1 = false;

                //Añadir los campos de "Acumulativos >"
                this.addColumnsOfCummulativeValuesToLeft(Dgv_ASTM_D95,numberOfRun);
                this.addColumnsOfCummulativeValuesToLeft(Dgv_ASTM_Single_Aperture, numberOfRun);
                this.addColumnsOfCummulativeValuesToLeft(Dgv_ASTM_D95_Accumulated_rigth_left,numberOfRun);
                this.addColumnsOfCummulativeValuesToLeft(Dgv_Single_Aperture_Accumulated_right_left, numberOfRun);

                //primera corrida 95% //Dgv_ASTM95_Detector_Number
                foreach (DataGridViewRow row in Dgv_ASTM95_Detector_Number.Rows)
                {
                    double accumulated = 0;
                    var cellvalue = row.Cells[3].Value;
                    int cellValuePlusOne = Convert.ToInt32(cellvalue) + 1;
                    //aumentar a la fila los valores acumulativos a la derecha (los que van arriba)
                    try
                    {
                        while (cellValuePlusOne > Convert.ToInt32(cellvalue))
                        {
                            accumulated = accumulated + Convert.ToDouble(Dgv_Particle_Data.Rows[cellValuePlusOne].Cells[2].Value);
                            cellValuePlusOne++;
                            if (accumulated > 100)
                            {
                                accumulated = 100;
                            }
                            Dgv_ASTM_D95.Rows[row.Index].Cells[5].Value = Math.Round(accumulated, 2);
                            Dgv_ASTM_D95_Accumulated_rigth_left.Rows[row.Index].Cells[5].Value = Math.Round(accumulated, 2);
                        }
                        double valor = 100 - Convert.ToDouble(Dgv_ASTM_D95_Accumulated_rigth_left.Rows[row.Index].Cells[2].Value);
                        Dgv_ASTM_D95_Accumulated_rigth_left.Rows[row.Index].Cells[5].Value = Math.Round(valor, 2);
                        Dgv_ASTM_D95.Rows[row.Index].Cells[5].Value = Math.Round(valor, 2);
                    }
                    catch (Exception ex)
                    {
                        Trace.WriteLine(ex.Message);
                    }
                }

                //primera corrida max%
                this.addCumulativeValuesToRightOfDataGridView(Dgv_Single_Aperture_Detector, Dgv_ASTM_Single_Aperture, 5);
                this.addCumulativeValuesToRightOfDataGridView(Dgv_Single_Aperture_Detector, Dgv_Single_Aperture_Accumulated_right_left, 5);

                Accumulated _accumulated = new Accumulated();

                //segunda Corrida 95%
                foreach (DataGridViewRow row2 in Dgv_ASTM95_Detector_Number.Rows)
                {
                    double acumarr2 = 0;
                    int q1 = Convert.ToInt32(row2.Cells[3].Value) + 1;
                    //aumentar a la fila los valores acumulativos a la derecha (los que van arriba)
                    try
                    {
                        while (q1 > Convert.ToInt32(row2.Cells[3].Value))
                        {
                            acumarr2 = acumarr2 + Convert.ToDouble(Dgv_Particle_Data.Rows[q1].Cells[3].Value);
                            q1++;
                            if (acumarr2 > 100)
                            {
                                acumarr2 = 100;
                            }
                            Dgv_ASTM_D95.Rows[row2.Index].Cells[6].Value = Math.Round(acumarr2, 2);
                            Dgv_ASTM_D95_Accumulated_rigth_left.Rows[row2.Index].Cells[6].Value = Math.Round(acumarr2, 2);
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
                //segunda Corrida max%
                this.addCumulativeValuesToRightOfDataGridView(Dgv_Single_Aperture_Detector, Dgv_ASTM_Single_Aperture, 6);
                this.addCumulativeValuesToRightOfDataGridView(Dgv_Single_Aperture_Detector, Dgv_Single_Aperture_Accumulated_right_left, 6);

                //tercera Corrida 95%
                this.addCumulativeValuesToRightOfDataGridView(Dgv_ASTM95_Detector_Number, Dgv_ASTM_D95_Accumulated_rigth_left, 7);
                this.addCumulativeValuesToRightOfDataGridView(Dgv_ASTM95_Detector_Number, Dgv_ASTM_D95, 7);

                //tercera Corrida max%
                this.addCumulativeValuesToRightOfDataGridView(Dgv_Single_Aperture_Detector, Dgv_ASTM_D95_Accumulated_rigth_left, 7);
                this.addCumulativeValuesToRightOfDataGridView(Dgv_Single_Aperture_Detector, Dgv_ASTM_D95, 7);

                this.addCumulativeValuesToRightOfDataGridView(Dgv_Single_Aperture_Detector, Dgv_ASTM_Single_Aperture, 7);
                this.addCumulativeValuesToRightOfDataGridView(Dgv_Single_Aperture_Detector, Dgv_Single_Aperture_Accumulated_right_left, 7);

                _accumulated.addCumulativeValuesToRight100(Dgv_ASTM_D95, Dgv_ASTM_D95_Accumulated_rigth_left);
                _accumulated.addCumulativeValuesToRight100(Dgv_ASTM_Single_Aperture, Dgv_Single_Aperture_Accumulated_right_left);
            }
            else if (Dgv_Particle_Data.Rows[0].Cells[3].Value.ToString() == "Run_2 (Vol%)")
            {
                numberOfRun = "2";
                //Añadir los campos de "Acumulativos <" 95%
                this.addColumnsOfCummulativeValues(Dgv_ASTM_D95, numberOfRun);
                this.addColumnsOfCummulativeValues(Dgv_ASTM_D95_Accumulated_rigth_left, numberOfRun);

                //Añadir los campos de "Acumulativos <" max%
                this.addColumnsOfCummulativeValues(Dgv_ASTM_Single_Aperture, numberOfRun);
                this.addColumnsOfCummulativeValues(Dgv_Single_Aperture_Accumulated_right_left, numberOfRun);

                Accumulated accumulated = new Accumulated(Dgv_Particle_Data);
         
                accumulated.addAccumulatedValuesToRightRunOne(Dgv_ASTM95_Detector_Number, Dgv_ASTM_D95);
                accumulated.addAccumulatedValuesToRightRunOne(Dgv_ASTM95_Detector_Number, Dgv_ASTM_D95_Accumulated_rigth_left);

                accumulated.addAccumulatedValuesToRightRunOne(Dgv_Single_Aperture_Detector, Dgv_ASTM_Single_Aperture);
                accumulated.addAccumulatedValuesToRightRunOne(Dgv_Single_Aperture_Detector, Dgv_Single_Aperture_Accumulated_right_left);

                accumulated.addAccumulatedValuesToRightRunTwo(Dgv_ASTM95_Detector_Number, Dgv_ASTM_D95);
                accumulated.addAccumulatedValuesToRightRunTwo(Dgv_ASTM95_Detector_Number, Dgv_ASTM_D95_Accumulated_rigth_left);

                accumulated.addAccumulatedValuesToRightRunTwo(Dgv_Single_Aperture_Detector, Dgv_ASTM_Single_Aperture);
                accumulated.addAccumulatedValuesToRightRunTwo(Dgv_Single_Aperture_Detector, Dgv_Single_Aperture_Accumulated_right_left);
                ch1 = false;

                this.addColumnsOfCummulativeValuesToLeft(Dgv_ASTM_D95, numberOfRun);
                this.addColumnsOfCummulativeValuesToLeft(Dgv_ASTM_D95_Accumulated_rigth_left, numberOfRun);

                this.addColumnsOfCummulativeValuesToLeft(Dgv_ASTM_Single_Aperture, numberOfRun);
                this.addColumnsOfCummulativeValuesToLeft(Dgv_Single_Aperture_Accumulated_right_left, numberOfRun);

                accumulated.addAccumulatedValuesToRightRunThree(Dgv_ASTM95_Detector_Number, Dgv_ASTM_D95,4);
                accumulated.addAccumulatedValuesToRightRunThree(Dgv_ASTM95_Detector_Number, Dgv_ASTM_D95_Accumulated_rigth_left,4);

                accumulated.addAccumulatedValuesToRightRunThree(Dgv_Single_Aperture_Detector, Dgv_ASTM_Single_Aperture,4);
                accumulated.addAccumulatedValuesToRightRunThree(Dgv_Single_Aperture_Detector, Dgv_Single_Aperture_Accumulated_right_left, 4);
  
                accumulated.addAccumulatedValuesToRightRunFor(Dgv_ASTM95_Detector_Number, Dgv_ASTM_D95);
                accumulated.addAccumulatedValuesToRightRunFor(Dgv_ASTM95_Detector_Number, Dgv_ASTM_D95_Accumulated_rigth_left);

                accumulated.addAccumulatedValuesToRightRunFor(Dgv_Single_Aperture_Detector, Dgv_ASTM_Single_Aperture);
                accumulated.addAccumulatedValuesToRightRunFor(Dgv_Single_Aperture_Detector, Dgv_Single_Aperture_Accumulated_right_left);

                accumulated.addCumulativeValuesToLeftBy100(Dgv_ASTM_D95, Dgv_ASTM_D95_Accumulated_rigth_left);
                accumulated.addCumulativeValuesToLeftBy100(Dgv_ASTM_Single_Aperture, Dgv_Single_Aperture_Accumulated_right_left);

            }
            else if (Dgv_Particle_Data.Rows[0].Cells[2].Value.ToString() == "Run_1 (Vol%)")
            {
                numberOfRun = "1";
                //Añadir los campos de "Acumulativos <" 95%
                this.addColumnsOfCummulativeValues(Dgv_ASTM_D95, numberOfRun);
                this.addColumnsOfCummulativeValues(Dgv_ASTM_D95_Accumulated_rigth_left, numberOfRun);

                //Añadir los campos de "Acumulativos <" max%
                this.addColumnsOfCummulativeValues(Dgv_ASTM_Single_Aperture, numberOfRun);
                this.addColumnsOfCummulativeValues(Dgv_Single_Aperture_Accumulated_right_left, numberOfRun);
                Accumulated accumulated = new Accumulated(Dgv_Particle_Data);

                //primera corrida 95%
                accumulated.addAccumulatedValuesToRightRunOne(Dgv_ASTM95_Detector_Number, Dgv_ASTM_D95);
                accumulated.addAccumulatedValuesToRightRunOne(Dgv_ASTM95_Detector_Number, Dgv_ASTM_D95_Accumulated_rigth_left);
                //primera corrida max%
                accumulated.addAccumulatedValuesToRightRunOne(Dgv_Single_Aperture_Detector, Dgv_ASTM_Single_Aperture);
                accumulated.addAccumulatedValuesToRightRunOne(Dgv_Single_Aperture_Detector, Dgv_Single_Aperture_Accumulated_right_left);
                ch1 = false;

                this.addColumnsOfCummulativeValuesToLeft(Dgv_ASTM_D95, numberOfRun);
                this.addColumnsOfCummulativeValuesToLeft(Dgv_ASTM_D95_Accumulated_rigth_left, numberOfRun);

                this.addColumnsOfCummulativeValuesToLeft(Dgv_ASTM_Single_Aperture, numberOfRun);
                this.addColumnsOfCummulativeValuesToLeft(Dgv_Single_Aperture_Accumulated_right_left, numberOfRun);

                //primera corrida 95%
                accumulated.addAccumulatedValuesToRightRunThree(Dgv_ASTM95_Detector_Number, Dgv_ASTM_D95, 3);
                accumulated.addAccumulatedValuesToRightRunThree(Dgv_ASTM95_Detector_Number, Dgv_ASTM_D95_Accumulated_rigth_left, 3);

                //primera corrida max%
                accumulated.addAccumulatedValuesToRightRunThree(Dgv_Single_Aperture_Detector, Dgv_ASTM_Single_Aperture, 3);
                accumulated.addAccumulatedValuesToRightRunThree(Dgv_Single_Aperture_Detector, Dgv_Single_Aperture_Accumulated_right_left, 3);

                corridas = 1;
                accumulated.addCumulativeValuesToLeftBy100RunOne(Dgv_ASTM_D95, Dgv_ASTM_D95_Accumulated_rigth_left);
                accumulated.addCumulativeValuesToLeftBy100RunOne(Dgv_ASTM_Single_Aperture, Dgv_Single_Aperture_Accumulated_right_left);
            }

            ch2 = false;

            //Aqui empieza el proceso del diferencial
            try
            {
                if (Dgv_Particle_Data.Rows[0].Cells[4].Value.ToString() == "Run_3 (Vol%)")
                {
                    // 3 corridas 
                    if (Dgv_ASTM_D95.Rows.Count == 2)
                    {
                        Dgv_ASTM95_Run_Differentials.Visible = true;
                        Dgv_Single_Aperture_Run_Differential.Visible = true;
                        //Funciones del diferencial 95%
                        Dgv_ASTM95_Run_Differentials.Rows.Add();
                        Dgv_ASTM95_Run_Differentials.Rows.Add();
                        Dgv_ASTM95_Run_Differentials.Rows.Add();

                        Dgv_ASTM_95_Differential.Rows.Add();
                        Dgv_ASTM_95_Differential.Rows.Add();
                        Dgv_ASTM_95_Differential.Rows.Add();
                        //Funciones del diferencial max%
                        Dgv_Single_Aperture_Run_Differential.Rows.Add();
                        Dgv_Single_Aperture_Run_Differential.Rows.Add();
                        Dgv_Single_Aperture_Run_Differential.Rows.Add();

                        Dgv_Single_Aperture_Differential.Rows.Add();
                        Dgv_Single_Aperture_Differential.Rows.Add();
                        Dgv_Single_Aperture_Differential.Rows.Add();
                        //Asignacion de variables de comparacion 95%
                        double comp1;
                        double val1 = 100000;
                        foreach (DataGridViewRow max1 in Dgv_ASTM_D95.Rows)
                        {
                            string max11 = Convert.ToString(max1.Cells[2].Value);
                            comp1 = Convert.ToDouble(max11);
                            if (comp1 < val1)
                            {
                                val1 = comp1;
                                name = max1.Cells[0].Value.ToString();
                                Dgv_ASTM95_Run_Differentials.Rows[0].Cells[0].Value = name;
                                Dgv_ASTM_95_Differential.Rows[0].Cells[0].Value = name;
                            }
                        }
                        Dgv_ASTM95_Run_Differentials.Rows[0].Cells[1].Value = Math.Round(val1, 2);
                        Dgv_ASTM_95_Differential.Rows[0].Cells[1].Value = Math.Round(val1, 2);

                        //Asignacion de variables de comparacion max%
                        double comp1z;
                        double val1z = 100000;
                        foreach (DataGridViewRow max1 in Dgv_ASTM_Single_Aperture.Rows)
                        {
                            string max11 = Convert.ToString(max1.Cells[2].Value);
                            comp1z = Convert.ToDouble(max11);
                            if (comp1z < val1z)
                            {
                                val1z = comp1z;
                                namez = max1.Cells[0].Value.ToString();
                                Dgv_Single_Aperture_Run_Differential.Rows[0].Cells[0].Value = namez;
                                Dgv_Single_Aperture_Differential.Rows[0].Cells[0].Value = namez;
                            }
                        }
                        Dgv_Single_Aperture_Run_Differential.Rows[0].Cells[1].Value = Math.Round(val1z, 2);
                        Dgv_Single_Aperture_Differential.Rows[0].Cells[1].Value = Math.Round(val1z, 2);

                        //2 95%
                        double comp2;
                        double val2 = 100000;
                        foreach (DataGridViewRow max2 in Dgv_ASTM_D95.Rows)
                        {
                            string max21 = Convert.ToString(max2.Cells[3].Value);
                            comp2 = Convert.ToDouble(max21);
                            if (comp2 < val2)
                            {
                                val2 = comp2;
                            }
                        }
                        Dgv_ASTM95_Run_Differentials.Rows[0].Cells[2].Value = Math.Round(val2, 2);
                        Dgv_ASTM_95_Differential.Rows[0].Cells[2].Value = Math.Round(val2, 2);
                        //2 max%
                        double comp2z;
                        double val2z = 100000;
                        foreach (DataGridViewRow max2 in Dgv_ASTM_Single_Aperture.Rows)
                        {
                            string max21 = Convert.ToString(max2.Cells[3].Value);
                            comp2z = Convert.ToDouble(max21);
                            if (comp2z < val2z)
                            {
                                val2z = comp2z;
                            }
                        }
                        Dgv_Single_Aperture_Run_Differential.Rows[0].Cells[2].Value = Math.Round(val2z, 2);
                        Dgv_Single_Aperture_Differential.Rows[0].Cells[2].Value = Math.Round(val2z, 2);

                        //3 95%
                        double comp3;
                        double val3 = 100000;

                        foreach (DataGridViewRow max3 in Dgv_ASTM_D95.Rows)
                        {
                            string max31 = Convert.ToString(max3.Cells[4].Value);
                            comp3 = Convert.ToDouble(max31);
                            if (comp3 < val3)
                            {
                                val3 = comp3;
                            }
                        }
                        Dgv_ASTM95_Run_Differentials.Rows[0].Cells[3].Value = Math.Round(val3, 2);
                        Dgv_ASTM_95_Differential.Rows[0].Cells[3].Value = Math.Round(val3, 2);
                        //3 max%
                        double comp3z;
                        double val3z = 100000;

                        foreach (DataGridViewRow max3 in Dgv_ASTM_Single_Aperture.Rows)
                        {
                            string max31 = Convert.ToString(max3.Cells[4].Value);
                            comp3z = Convert.ToDouble(max31);
                            if (comp3z < val3z)
                            {
                                val3z = comp3z;
                            }
                        }
                        Dgv_Single_Aperture_Run_Differential.Rows[0].Cells[3].Value = Math.Round(val3z, 2);
                        Dgv_Single_Aperture_Differential.Rows[0].Cells[3].Value = Math.Round(val3z, 2);

                        //4 95%
                        double comp4;
                        double val4 = 1000000;
                        string name2;
                        foreach (DataGridViewRow max4 in Dgv_ASTM_D95.Rows)
                        {
                            string max41 = Convert.ToString(max4.Cells[5].Value);
                            comp4 = Convert.ToDouble(max41);
                            if (comp4 < val4)
                            {
                                val4 = comp4;
                                name2 = max4.Cells[0].Value.ToString();
                                Dgv_ASTM95_Run_Differentials.Rows[2].Cells[0].Value = name2;
                                Dgv_ASTM_95_Differential.Rows[2].Cells[0].Value = name2;
                            }
                        }
                        Dgv_ASTM95_Run_Differentials.Rows[2].Cells[1].Value = Math.Round(val4, 2);
                        Dgv_ASTM_95_Differential.Rows[2].Cells[1].Value = Math.Round(val4, 2);
                        //4 max%
                        double comp4z;
                        double val4z = 1000000;
                        string name2z;
                        foreach (DataGridViewRow max4 in Dgv_ASTM_Single_Aperture.Rows)
                        {
                            string max41 = Convert.ToString(max4.Cells[5].Value);
                            comp4z = Convert.ToDouble(max41);
                            if (comp4z < val4z)
                            {
                                val4z = comp4z;
                                name2z = max4.Cells[0].Value.ToString();
                                Dgv_Single_Aperture_Run_Differential.Rows[2].Cells[0].Value = name2z;
                                Dgv_Single_Aperture_Differential.Rows[2].Cells[0].Value = name2z;
                            }
                        }
                        Dgv_Single_Aperture_Run_Differential.Rows[2].Cells[1].Value = Math.Round(val4z, 2);
                        Dgv_Single_Aperture_Differential.Rows[2].Cells[1].Value = Math.Round(val4z, 2);

                        //5 95%
                        double comp5;
                        double val5 = 1000000;
                        foreach (DataGridViewRow max5 in Dgv_ASTM_D95.Rows)
                        {
                            string max51 = Convert.ToString(max5.Cells[6].Value);
                            comp5 = Convert.ToDouble(max51);
                            if (comp5 < val5)
                            {
                                val5 = comp5;
                            }
                        }
                        Dgv_ASTM95_Run_Differentials.Rows[2].Cells[2].Value = Math.Round(val5, 2);
                        Dgv_ASTM_95_Differential.Rows[2].Cells[2].Value = Math.Round(val5, 2);
                        //5 max%
                        double comp5z;
                        double val5z = 1000000;
                        foreach (DataGridViewRow max5 in Dgv_ASTM_Single_Aperture.Rows)
                        {
                            string max51 = Convert.ToString(max5.Cells[6].Value);
                            comp5z = Convert.ToDouble(max51);
                            if (comp5z < val5z)
                            {
                                val5z = comp5z;
                            }
                        }
                        Dgv_Single_Aperture_Run_Differential.Rows[2].Cells[2].Value = Math.Round(val5z, 2);
                        Dgv_Single_Aperture_Differential.Rows[2].Cells[2].Value = Math.Round(val5z, 2);

                        //6 95%
                        double comp6;
                        double val6 = 10000;
                        foreach (DataGridViewRow max6 in Dgv_ASTM_D95.Rows)
                        {
                            string max61 = Convert.ToString(max6.Cells[7].Value);
                            comp6 = Convert.ToDouble(max61);
                            if (comp6 < val6)
                            {
                                val6 = comp6;
                            }
                        }
                        Dgv_ASTM95_Run_Differentials.Rows[2].Cells[3].Value = Math.Round(val6, 2);
                        Dgv_ASTM_95_Differential.Rows[2].Cells[3].Value = Math.Round(val6, 2);
                        //6 max%
                        double comp6z;
                        double val6z = 10000;
                        foreach (DataGridViewRow max6 in Dgv_ASTM_Single_Aperture.Rows)
                        {
                            string max61 = Convert.ToString(max6.Cells[7].Value);
                            comp6z = Convert.ToDouble(max61);
                            if (comp6z < val6z)
                            {
                                val6z = comp6z;
                            }
                        }
                        Dgv_Single_Aperture_Run_Differential.Rows[2].Cells[3].Value = Math.Round(val6z, 2);
                        Dgv_Single_Aperture_Differential.Rows[2].Cells[3].Value = Math.Round(val6z, 2);

                        //Crear el diferencial 95%
                        Dgv_ASTM95_Run_Differentials.Rows[1].Cells[1].Value =
                        Math.Round(Convert.ToDouble(100 - (Convert.ToDouble(Dgv_ASTM95_Run_Differentials.Rows[2].Cells[1].Value) +
                        Convert.ToDouble(Dgv_ASTM95_Run_Differentials.Rows[0].Cells[1].Value))), 2);

                        Dgv_ASTM95_Run_Differentials.Rows[1].Cells[2].Value =
                        Math.Round(Convert.ToDouble(100 - (Convert.ToDouble(Dgv_ASTM95_Run_Differentials.Rows[2].Cells[2].Value) +
                        Convert.ToDouble(Dgv_ASTM95_Run_Differentials.Rows[0].Cells[2].Value))), 2);

                        Dgv_ASTM95_Run_Differentials.Rows[1].Cells[3].Value =
                        Math.Round(Convert.ToDouble(100 - (Convert.ToDouble(Dgv_ASTM95_Run_Differentials.Rows[2].Cells[3].Value) +
                        Convert.ToDouble(Dgv_ASTM95_Run_Differentials.Rows[0].Cells[3].Value))), 2);

                        Dgv_ASTM_95_Differential.Rows[1].Cells[1].Value =
                        Math.Round(Convert.ToDouble(100 - (Convert.ToDouble(Dgv_ASTM95_Run_Differentials.Rows[2].Cells[1].Value) +
                        Convert.ToDouble(Dgv_ASTM95_Run_Differentials.Rows[0].Cells[1].Value))), 2);

                        Dgv_ASTM_95_Differential.Rows[1].Cells[2].Value =
                        Math.Round(Convert.ToDouble(100 - (Convert.ToDouble(Dgv_ASTM95_Run_Differentials.Rows[2].Cells[2].Value) +
                        Convert.ToDouble(Dgv_ASTM95_Run_Differentials.Rows[0].Cells[2].Value))), 2);

                        Dgv_ASTM_95_Differential.Rows[1].Cells[3].Value =
                        Math.Round(Convert.ToDouble(100 - (Convert.ToDouble(Dgv_ASTM95_Run_Differentials.Rows[2].Cells[3].Value) +
                        Convert.ToDouble(Dgv_ASTM95_Run_Differentials.Rows[0].Cells[3].Value))), 2);

                        //Crear el diferencial max%
                        Dgv_Single_Aperture_Run_Differential.Rows[1].Cells[1].Value =
                        Math.Round(Convert.ToDouble(100 - (Convert.ToDouble(Dgv_Single_Aperture_Run_Differential.Rows[2].Cells[1].Value) +
                        Convert.ToDouble(Dgv_Single_Aperture_Run_Differential.Rows[0].Cells[1].Value))), 2);

                        Dgv_Single_Aperture_Run_Differential.Rows[1].Cells[2].Value =
                        Math.Round(Convert.ToDouble(100 - (Convert.ToDouble(Dgv_Single_Aperture_Run_Differential.Rows[2].Cells[2].Value) +
                        Convert.ToDouble(Dgv_Single_Aperture_Run_Differential.Rows[0].Cells[2].Value))), 2);

                        Dgv_Single_Aperture_Run_Differential.Rows[1].Cells[3].Value =
                        Math.Round(Convert.ToDouble(100 - (Convert.ToDouble(Dgv_Single_Aperture_Run_Differential.Rows[2].Cells[3].Value) +
                        Convert.ToDouble(Dgv_Single_Aperture_Run_Differential.Rows[0].Cells[3].Value))), 2);

                        Dgv_Single_Aperture_Differential.Rows[1].Cells[1].Value =
                        Math.Round(Convert.ToDouble(100 - (Convert.ToDouble(Dgv_Single_Aperture_Run_Differential.Rows[2].Cells[1].Value) +
                        Convert.ToDouble(Dgv_Single_Aperture_Run_Differential.Rows[0].Cells[1].Value))), 2);

                        Dgv_Single_Aperture_Differential.Rows[1].Cells[2].Value =
                        Math.Round(Convert.ToDouble(100 - (Convert.ToDouble(Dgv_Single_Aperture_Run_Differential.Rows[2].Cells[2].Value) +
                        Convert.ToDouble(Dgv_Single_Aperture_Run_Differential.Rows[0].Cells[2].Value))), 2);

                        Dgv_Single_Aperture_Differential.Rows[1].Cells[3].Value =
                        Math.Round(Convert.ToDouble(100 - (Convert.ToDouble(Dgv_Single_Aperture_Run_Differential.Rows[2].Cells[3].Value) +
                        Convert.ToDouble(Dgv_Single_Aperture_Run_Differential.Rows[0].Cells[3].Value))), 2);

                        renombrar();
                    }
                    else if (Dgv_ASTM_D95.Rows.Count > 2)
                    {
                        Dgv_ASTM95_Run_Differentials.Visible = true;
                        Dgv_Single_Aperture_Run_Differential.Visible = true;

                        //Para 95%
                        int n_filas = Convert.ToInt32(Dgv_ASTM_D95.Rows.Count.ToString()) + 1;
                        int contador = 1;
                        while (contador < n_filas)
                        {
                            Dgv_ASTM95_Run_Differentials.Rows.Add();
                            Dgv_ASTM_95_Differential.Rows.Add();
                            contador++;
                        }
                        //Para max%
                        int n_filas1 = Convert.ToInt32(Dgv_ASTM_Single_Aperture.Rows.Count.ToString()) + 1;
                        int contador1 = 1;
                        while (contador1 < n_filas1)
                        {
                            Dgv_Single_Aperture_Run_Differential.Rows.Add();
                            Dgv_Single_Aperture_Differential.Rows.Add();
                            contador1++;
                        }

                        //llenar la primera y la ultima columna 95%
                        double comp;
                        double val = 100000;
                        foreach (DataGridViewRow max in Dgv_ASTM_D95.Rows)
                        {
                            string max1 = Convert.ToString(max.Cells[2].Value);
                            comp = Convert.ToDouble(max1);
                            if (comp < val)
                            {
                                val = comp;
                                name = max.Cells[0].Value.ToString();
                                Dgv_ASTM95_Run_Differentials.Rows[0].Cells[0].Value = name;
                                Dgv_ASTM_95_Differential.Rows[0].Cells[0].Value = name;
                            }
                        }
                        Dgv_ASTM95_Run_Differentials.Rows[0].Cells[1].Value = Math.Round(val, 2);
                        Dgv_ASTM_95_Differential.Rows[0].Cells[1].Value = Math.Round(val, 2);

                        double comp1;
                        double val1 = 100000;
                        foreach (DataGridViewRow max1 in Dgv_ASTM_D95.Rows)
                        {
                            string max11 = Convert.ToString(max1.Cells[5].Value);
                            comp1 = Convert.ToDouble(max11);
                            if (comp1 < val1)
                            {
                                val1 = comp1;
                            }
                        }
                        Dgv_ASTM95_Run_Differentials.Rows[n_filas - 1].Cells[1].Value = Math.Round(val1, 2);
                        Dgv_ASTM_95_Differential.Rows[n_filas - 1].Cells[1].Value = Math.Round(val1, 2);
                        //llenar la primera y la ultima columna max%
                        double comp1z;
                        double val1z = 100000;
                        foreach (DataGridViewRow max in Dgv_ASTM_Single_Aperture.Rows)
                        {
                            string max1 = Convert.ToString(max.Cells[2].Value);
                            comp1z = Convert.ToDouble(max1);
                            if (comp1z < val1z)
                            {
                                val1z = comp1z;
                                namez = max.Cells[0].Value.ToString();
                                Dgv_Single_Aperture_Run_Differential.Rows[0].Cells[0].Value = namez;
                                Dgv_Single_Aperture_Differential.Rows[0].Cells[0].Value = namez;
                            }
                        }
                        Dgv_Single_Aperture_Run_Differential.Rows[0].Cells[1].Value = Math.Round(val1z, 2);
                        Dgv_Single_Aperture_Differential.Rows[0].Cells[1].Value = Math.Round(val1z, 2);

                        double comp1z1;
                        double val1z1 = 100000;
                        foreach (DataGridViewRow max1 in Dgv_ASTM_Single_Aperture.Rows)
                        {
                            string max11 = Convert.ToString(max1.Cells[5].Value);
                            comp1z1 = Convert.ToDouble(max11);
                            if (comp1z1 < val1z1)
                            {
                                val1z1 = comp1z1;
                            }
                        }
                        Dgv_Single_Aperture_Run_Differential.Rows[n_filas1 - 1].Cells[1].Value = Math.Round(val1z1, 2);
                        Dgv_Single_Aperture_Differential.Rows[n_filas1 - 1].Cells[1].Value = Math.Round(val1z1, 2);

                        //segunda carrera 95%
                        double comp2;
                        double val2 = 100000;
                        foreach (DataGridViewRow max21 in Dgv_ASTM_D95.Rows)
                        {
                            string max2 = Convert.ToString(max21.Cells[3].Value);
                            comp2 = Convert.ToDouble(max2);
                            if (comp2 < val2)
                            {
                                val2 = comp2;
                            }
                        }
                        Dgv_ASTM95_Run_Differentials.Rows[0].Cells[2].Value = Math.Round(val2, 2);
                        Dgv_ASTM_95_Differential.Rows[0].Cells[2].Value = Math.Round(val2, 2);

                        double comp4;
                        double val4 = 100000;
                        foreach (DataGridViewRow max4 in Dgv_ASTM_D95.Rows)
                        {
                            string max41 = Convert.ToString(max4.Cells[6].Value);
                            comp4 = Convert.ToDouble(max41);
                            if (comp4 < val4)
                            {
                                val4 = comp4;
                                lname = max4.Cells[0].Value.ToString();
                                Dgv_ASTM95_Run_Differentials.Rows[n_filas - 1].Cells[0].Value = lname;
                                Dgv_ASTM_95_Differential.Rows[n_filas - 1].Cells[0].Value = lname;
                            }
                        }
                        Dgv_ASTM95_Run_Differentials.Rows[n_filas - 1].Cells[2].Value = Math.Round(val4, 2);
                        Dgv_ASTM_95_Differential.Rows[n_filas - 1].Cells[2].Value = Math.Round(val4, 2);
                        //segunda carrera max%
                        double comp2z;
                        double val2z = 100000;
                        foreach (DataGridViewRow max21 in Dgv_ASTM_Single_Aperture.Rows)
                        {
                            string max2 = Convert.ToString(max21.Cells[3].Value);
                            comp2z = Convert.ToDouble(max2);
                            if (comp2z < val2z)
                            {
                                val2z = comp2z;
                            }
                        }
                        Dgv_Single_Aperture_Run_Differential.Rows[0].Cells[2].Value = Math.Round(val2z, 2);
                        Dgv_Single_Aperture_Differential.Rows[0].Cells[2].Value = Math.Round(val2z, 2);

                        double comp4z;
                        double val4z = 100000;
                        foreach (DataGridViewRow max4 in Dgv_ASTM_Single_Aperture.Rows)
                        {
                            string max41 = Convert.ToString(max4.Cells[6].Value);
                            comp4z = Convert.ToDouble(max41);
                            if (comp4z < val4z)
                            {
                                val4z = comp4z;
                                lnamez = max4.Cells[0].Value.ToString();
                                Dgv_Single_Aperture_Run_Differential.Rows[n_filas1 - 1].Cells[0].Value = lnamez;
                                Dgv_Single_Aperture_Differential.Rows[n_filas1 - 1].Cells[0].Value = lnamez;
                            }
                        }
                        Dgv_Single_Aperture_Run_Differential.Rows[n_filas1 - 1].Cells[2].Value = Math.Round(val4z, 2);
                        Dgv_Single_Aperture_Differential.Rows[n_filas1 - 1].Cells[2].Value = Math.Round(val4z, 2);

                        //tercera carrera 95%
                        double comp5;
                        double val5 = 100000;
                        foreach (DataGridViewRow max51 in Dgv_ASTM_D95.Rows)
                        {
                            string max5 = Convert.ToString(max51.Cells[4].Value);
                            comp5 = Convert.ToDouble(max5);
                            if (comp5 < val5)
                            {
                                val5 = comp5;
                            }
                        }
                        Dgv_ASTM95_Run_Differentials.Rows[0].Cells[3].Value = Math.Round(val5, 2);
                        Dgv_ASTM_95_Differential.Rows[0].Cells[3].Value = Math.Round(val5, 2);

                        double comp6;
                        double val6 = 100000;
                        foreach (DataGridViewRow max6 in Dgv_ASTM_D95.Rows)
                        {
                            string max61 = Convert.ToString(max6.Cells[7].Value);
                            comp6 = Convert.ToDouble(max61);
                            if (comp6 < val6)
                            {
                                val6 = comp6;
                            }
                        }
                        Dgv_ASTM95_Run_Differentials.Rows[n_filas - 1].Cells[3].Value = Math.Round(val6, 2);
                        Dgv_ASTM_95_Differential.Rows[n_filas - 1].Cells[3].Value = Math.Round(val6, 2);
                        //tercera carrera max%
                        double comp5z;
                        double val5z = 100000;
                        foreach (DataGridViewRow max51 in Dgv_ASTM_Single_Aperture.Rows)
                        {
                            string max5 = Convert.ToString(max51.Cells[4].Value);
                            comp5z = Convert.ToDouble(max5);
                            if (comp5z < val5z)
                            {
                                val5z = comp5z;
                            }
                        }
                        Dgv_Single_Aperture_Run_Differential.Rows[0].Cells[3].Value = Math.Round(val5z, 2);
                        Dgv_Single_Aperture_Differential.Rows[0].Cells[3].Value = Math.Round(val5z, 2);

                        double comp6z;
                        double val6z = 100000;
                        foreach (DataGridViewRow max6 in Dgv_ASTM_Single_Aperture.Rows)
                        {
                            string max61 = Convert.ToString(max6.Cells[7].Value);
                            comp6z = Convert.ToDouble(max61);
                            if (comp6z < val6z)
                            {
                                val6z = comp6z;
                            }
                        }
                        Dgv_Single_Aperture_Run_Differential.Rows[n_filas1 - 1].Cells[3].Value = Math.Round(val6z, 2);
                        Dgv_Single_Aperture_Differential.Rows[n_filas1 - 1].Cells[3].Value = Math.Round(val6z, 2);

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
                        foreach (DataGridViewRow con in Dgv_ASTM_D95.Rows)
                        {
                            //Primera Corrida
                            while (com < 1)
                            {
                                acumulador = Convert.ToDouble(Dgv_ASTM95_Run_Differentials.Rows[0].Cells[1].Value);
                                com++;
                            }
                            if (con.Cells[0].Value.ToString() != name.ToString())
                            {
                                //Operaciones
                                calculo = Convert.ToDouble(con.Cells[2].Value) - Convert.ToDouble(Dgv_ASTM_D95.Rows[check1].Cells[2].Value.ToString());
                                Dgv_ASTM95_Run_Differentials.Rows[check].Cells[1].Value = Math.Round(calculo, 2);
                                Dgv_ASTM_95_Differential.Rows[check].Cells[1].Value = Math.Round(calculo, 2);
                                //Aumento de acumulador
                                acumulador = acumulador + Convert.ToDouble(con.Cells[2].Value.ToString());
                            }
                            //Segunda Corrida
                            while (com1 < 1)
                            {
                                acumulador1 = Convert.ToDouble(Dgv_ASTM95_Run_Differentials.Rows[0].Cells[2].Value);
                                com1++;
                            }
                            if (con.Cells[0].Value.ToString() != name.ToString())
                            {
                                //Operaciones
                                calculo1 = ((Convert.ToDouble(con.Cells[3].Value)) - (Convert.ToDouble(Dgv_ASTM_D95.Rows[check1].Cells[3].Value.ToString())));
                                Dgv_ASTM95_Run_Differentials.Rows[check].Cells[2].Value = Math.Round(calculo1, 2);
                                Dgv_ASTM_95_Differential.Rows[check].Cells[2].Value = Math.Round(calculo1, 2);

                                //Aumento de acumulador
                                acumulador1 = acumulador1 + Convert.ToDouble(con.Cells[3].Value.ToString());
                            }
                            //Tercera Corrida
                            while (com2 < 1)
                            {
                                acumulador2 = Convert.ToDouble(Dgv_ASTM95_Run_Differentials.Rows[0].Cells[3].Value);
                                com2++;
                            }
                            if (con.Cells[0].Value.ToString() != name.ToString())
                            {
                                //Operaciones
                                calculo2 = ((Convert.ToDouble(con.Cells[4].Value)) - (Convert.ToDouble(Dgv_ASTM_D95.Rows[check1].Cells[4].Value.ToString())));
                                Dgv_ASTM95_Run_Differentials.Rows[check].Cells[3].Value = Math.Round(calculo2, 2);
                                Dgv_ASTM_95_Differential.Rows[check].Cells[3].Value = Math.Round(calculo2, 2);

                                //Aumento de acumulador
                                acumulador2 = acumulador2 + Convert.ToDouble(con.Cells[4].Value.ToString());
                                check++;
                                check1++;
                            }
                        }
                        //Llenar los ultimos datos max%
                        foreach (DataGridViewRow con in Dgv_ASTM_Single_Aperture.Rows)
                        {
                            //Primera Corrida
                            while (comz < 1)
                            {
                                acumuladorz = Convert.ToDouble(Dgv_Single_Aperture_Run_Differential.Rows[0].Cells[1].Value);
                                comz++;
                            }
                            if (con.Cells[0].Value.ToString() != namez.ToString())
                            {
                                //Operaciones
                                calculoz = Convert.ToDouble(con.Cells[2].Value) - Convert.ToDouble(Dgv_ASTM_Single_Aperture.Rows[check1z].Cells[2].Value.ToString());
                                Dgv_Single_Aperture_Run_Differential.Rows[checkz].Cells[1].Value = Math.Round(calculoz, 2);
                                Dgv_Single_Aperture_Differential.Rows[checkz].Cells[1].Value = Math.Round(calculoz, 2);
                                //Aumento de acumulador
                                acumuladorz = acumuladorz + Convert.ToDouble(con.Cells[2].Value.ToString());
                            }
                            //Segunda Corrida
                            while (com1z < 1)
                            {
                                acumulador1z = Convert.ToDouble(Dgv_Single_Aperture_Run_Differential.Rows[0].Cells[2].Value);
                                com1z++;
                            }
                            if (con.Cells[0].Value.ToString() != namez.ToString())
                            {
                                //Operaciones
                                calculo1z = ((Convert.ToDouble(con.Cells[3].Value)) - (Convert.ToDouble(Dgv_ASTM_Single_Aperture.Rows[check1z].Cells[3].Value.ToString())));
                                Dgv_Single_Aperture_Run_Differential.Rows[checkz].Cells[2].Value = Math.Round(calculo1z, 2);
                                Dgv_Single_Aperture_Differential.Rows[checkz].Cells[2].Value = Math.Round(calculo1z, 2);

                                //Aumento de acumulador
                                acumulador1z = acumulador1z + Convert.ToDouble(con.Cells[3].Value.ToString());
                            }
                            //Tercera Corrida
                            while (com2z < 1)
                            {
                                acumulador2z = Convert.ToDouble(Dgv_Single_Aperture_Run_Differential.Rows[0].Cells[3].Value);
                                com2z++;
                            }
                            if (con.Cells[0].Value.ToString() != namez.ToString())
                            {
                                //Operaciones
                                calculo2z = ((Convert.ToDouble(con.Cells[4].Value)) - (Convert.ToDouble(Dgv_ASTM_Single_Aperture.Rows[check1z].Cells[4].Value.ToString())));
                                Dgv_Single_Aperture_Run_Differential.Rows[checkz].Cells[3].Value = Math.Round(calculo2z, 2);
                                Dgv_Single_Aperture_Differential.Rows[checkz].Cells[3].Value = Math.Round(calculo2z, 2);

                                //Aumento de acumulador
                                acumulador2z = acumulador2z + Convert.ToDouble(con.Cells[4].Value.ToString());
                                checkz++;
                                check1z++;
                            }
                        }
                        renombrar1();
                    }
                }
                else if (Dgv_Particle_Data.Rows[0].Cells[3].Value.ToString() == "Run_2 (Vol%)")
                {
                    //2 corridas
                    if (Dgv_ASTM_D95.Rows.Count == 2)
                    {
                        Dgv_ASTM95_Run_Differentials.Visible = true;
                        Dgv_Single_Aperture_Run_Differential.Visible = true;
                        //Funciones del diferencial 95%
                        Dgv_ASTM95_Run_Differentials.Rows.Add();
                        Dgv_ASTM95_Run_Differentials.Rows.Add();
                        Dgv_ASTM95_Run_Differentials.Rows.Add();

                        Dgv_ASTM_95_Differential.Rows.Add();
                        Dgv_ASTM_95_Differential.Rows.Add();
                        Dgv_ASTM_95_Differential.Rows.Add();
                        //Funciones del diferencial max%
                        Dgv_Single_Aperture_Run_Differential.Rows.Add();
                        Dgv_Single_Aperture_Run_Differential.Rows.Add();
                        Dgv_Single_Aperture_Run_Differential.Rows.Add();

                        Dgv_Single_Aperture_Differential.Rows.Add();
                        Dgv_Single_Aperture_Differential.Rows.Add();
                        Dgv_Single_Aperture_Differential.Rows.Add();

                        //Asignacion de variables de comparacion 95%
                        double comp1;
                        double val1 = 100000;
                        foreach (DataGridViewRow max1 in Dgv_ASTM_D95.Rows)
                        {
                            string max11 = Convert.ToString(max1.Cells[2].Value);
                            comp1 = Convert.ToDouble(max11);
                            if (comp1 < val1)
                            {
                                val1 = comp1;
                                name = max1.Cells[0].Value.ToString();
                                Dgv_ASTM95_Run_Differentials.Rows[0].Cells[0].Value = name;
                                Dgv_ASTM_95_Differential.Rows[0].Cells[0].Value = name;
                            }
                        }
                        Dgv_ASTM95_Run_Differentials.Rows[0].Cells[1].Value = Math.Round(val1, 2);
                        Dgv_ASTM_95_Differential.Rows[0].Cells[1].Value = Math.Round(val1, 2);
                        //Asignacion de variables de comparacion max%
                        double comp1z;
                        double val1z = 100000;
                        foreach (DataGridViewRow max1 in Dgv_ASTM_Single_Aperture.Rows)
                        {
                            string max11 = Convert.ToString(max1.Cells[2].Value);
                            comp1z = Convert.ToDouble(max11);
                            if (comp1z < val1z)
                            {
                                val1z = comp1z;
                                namez = max1.Cells[0].Value.ToString();
                                Dgv_Single_Aperture_Run_Differential.Rows[0].Cells[0].Value = namez;
                                Dgv_Single_Aperture_Differential.Rows[0].Cells[0].Value = namez;
                            }
                        }
                        Dgv_Single_Aperture_Run_Differential.Rows[0].Cells[1].Value = Math.Round(val1z, 2);
                        Dgv_Single_Aperture_Differential.Rows[0].Cells[1].Value = Math.Round(val1z, 2);

                        //2 95%
                        double comp2;
                        double val2 = 100000;
                        foreach (DataGridViewRow max2 in Dgv_ASTM_D95.Rows)
                        {
                            string max21 = Convert.ToString(max2.Cells[3].Value);
                            comp2 = Convert.ToDouble(max21);
                            if (comp2 < val2)
                            {
                                val2 = comp2;
                            }
                        }
                        Dgv_ASTM95_Run_Differentials.Rows[0].Cells[2].Value = Math.Round(val2, 2);
                        Dgv_ASTM_95_Differential.Rows[0].Cells[2].Value = Math.Round(val2, 2);
                        //2 95%
                        double comp2z;
                        double val2z = 100000;
                        foreach (DataGridViewRow max2 in Dgv_ASTM_Single_Aperture.Rows)
                        {
                            string max21 = Convert.ToString(max2.Cells[3].Value);
                            comp2z = Convert.ToDouble(max21);
                            if (comp2z < val2z)
                            {
                                val2z = comp2z;
                            }
                        }
                        Dgv_Single_Aperture_Run_Differential.Rows[0].Cells[2].Value = Math.Round(val2z, 2);
                        Dgv_Single_Aperture_Differential.Rows[0].Cells[2].Value = Math.Round(val2z, 2);

                        //4 95%
                        double comp4;
                        double val4 = 1000000;
                        string name2;
                        foreach (DataGridViewRow max4 in Dgv_ASTM_D95.Rows)
                        {
                            string max41 = Convert.ToString(max4.Cells[4].Value);
                            comp4 = Convert.ToDouble(max41);
                            if (comp4 < val4)
                            {
                                val4 = comp4;
                                name2 = max4.Cells[0].Value.ToString();
                                Dgv_ASTM95_Run_Differentials.Rows[2].Cells[0].Value = name2;
                                Dgv_ASTM_95_Differential.Rows[2].Cells[0].Value = name2;
                            }
                        }
                        Dgv_ASTM95_Run_Differentials.Rows[2].Cells[1].Value = Math.Round(val4, 2);
                        Dgv_ASTM_95_Differential.Rows[2].Cells[1].Value = Math.Round(val4, 2);
                        //4 max%
                        double comp4z;
                        double val4z = 1000000;
                        string name2z;
                        foreach (DataGridViewRow max4 in Dgv_ASTM_Single_Aperture.Rows)
                        {
                            string max41 = Convert.ToString(max4.Cells[4].Value);
                            comp4z = Convert.ToDouble(max41);
                            if (comp4z < val4z)
                            {
                                val4z = comp4z;
                                name2z = max4.Cells[0].Value.ToString();
                                Dgv_Single_Aperture_Run_Differential.Rows[2].Cells[0].Value = name2z;
                                Dgv_Single_Aperture_Differential.Rows[2].Cells[0].Value = name2z;
                            }
                        }
                        Dgv_Single_Aperture_Run_Differential.Rows[2].Cells[1].Value = Math.Round(val4z, 2);
                        Dgv_Single_Aperture_Differential.Rows[2].Cells[1].Value = Math.Round(val4z, 2);

                        //5 95%
                        double comp5;
                        double val5 = 1000000;
                        foreach (DataGridViewRow max5 in Dgv_ASTM_D95.Rows)
                        {
                            string max51 = Convert.ToString(max5.Cells[5].Value);
                            comp5 = Convert.ToDouble(max51);
                            if (comp5 < val5)
                            {
                                val5 = comp5;
                            }
                        }
                        Dgv_ASTM95_Run_Differentials.Rows[2].Cells[2].Value = Math.Round(val5, 2);
                        Dgv_ASTM_95_Differential.Rows[2].Cells[2].Value = Math.Round(val5, 2);
                        //5 max%
                        double comp5z;
                        double val5z = 1000000;
                        foreach (DataGridViewRow max5 in Dgv_ASTM_Single_Aperture.Rows)
                        {
                            string max51 = Convert.ToString(max5.Cells[5].Value);
                            comp5z = Convert.ToDouble(max51);
                            if (comp5z < val5z)
                            {
                                val5z = comp5z;
                            }
                        }
                        Dgv_Single_Aperture_Run_Differential.Rows[2].Cells[2].Value = Math.Round(val5z, 2);
                        Dgv_Single_Aperture_Differential.Rows[2].Cells[2].Value = Math.Round(val5z, 2);

                        //Crear el diferencial 95%
                        Dgv_ASTM95_Run_Differentials.Rows[1].Cells[1].Value =
                        Math.Round(Convert.ToDouble(100 - (Convert.ToDouble(Dgv_ASTM95_Run_Differentials.Rows[2].Cells[1].Value) +
                        Convert.ToDouble(Dgv_ASTM95_Run_Differentials.Rows[0].Cells[1].Value))), 2);

                        Dgv_ASTM95_Run_Differentials.Rows[1].Cells[2].Value =
                        Math.Round(Convert.ToDouble(100 - (Convert.ToDouble(Dgv_ASTM95_Run_Differentials.Rows[2].Cells[2].Value) +
                        Convert.ToDouble(Dgv_ASTM95_Run_Differentials.Rows[0].Cells[2].Value))), 2);

                        Dgv_ASTM_95_Differential.Rows[1].Cells[1].Value =
                        Math.Round(Convert.ToDouble(100 - (Convert.ToDouble(Dgv_ASTM95_Run_Differentials.Rows[2].Cells[1].Value) +
                        Convert.ToDouble(Dgv_ASTM95_Run_Differentials.Rows[0].Cells[1].Value))), 2);

                        Dgv_ASTM_95_Differential.Rows[1].Cells[2].Value =
                        Math.Round(Convert.ToDouble(100 - (Convert.ToDouble(Dgv_ASTM95_Run_Differentials.Rows[2].Cells[2].Value) +
                        Convert.ToDouble(Dgv_ASTM95_Run_Differentials.Rows[0].Cells[2].Value))), 2);
        
                        //Crear el diferencial max%
                        Dgv_Single_Aperture_Run_Differential.Rows[1].Cells[1].Value =
                        Math.Round(Convert.ToDouble(100 - (Convert.ToDouble(Dgv_Single_Aperture_Run_Differential.Rows[2].Cells[1].Value) +
                        Convert.ToDouble(Dgv_Single_Aperture_Run_Differential.Rows[0].Cells[1].Value))), 2);

                        Dgv_Single_Aperture_Run_Differential.Rows[1].Cells[2].Value =
                        Math.Round(Convert.ToDouble(100 - (Convert.ToDouble(Dgv_Single_Aperture_Run_Differential.Rows[2].Cells[2].Value) +
                        Convert.ToDouble(Dgv_Single_Aperture_Run_Differential.Rows[0].Cells[2].Value))), 2);

                        Dgv_Single_Aperture_Differential.Rows[1].Cells[1].Value =
                        Math.Round(Convert.ToDouble(100 - (Convert.ToDouble(Dgv_Single_Aperture_Run_Differential.Rows[2].Cells[1].Value) +
                        Convert.ToDouble(Dgv_Single_Aperture_Run_Differential.Rows[0].Cells[1].Value))), 2);

                        Dgv_Single_Aperture_Differential.Rows[1].Cells[2].Value =
                        Math.Round(Convert.ToDouble(100 - (Convert.ToDouble(Dgv_Single_Aperture_Run_Differential.Rows[2].Cells[2].Value) +
                        Convert.ToDouble(Dgv_Single_Aperture_Run_Differential.Rows[0].Cells[2].Value))), 2);

                        renombrar();
                    }
                    else if (Dgv_ASTM_D95.Rows.Count > 2)
                    {
                        Dgv_ASTM95_Run_Differentials.Visible = true;
                        Dgv_Single_Aperture_Run_Differential.Visible = true;

                        //95%
                        int n_filas = Convert.ToInt32(Dgv_ASTM_D95.Rows.Count.ToString()) + 1;
                        int contador = 1;
                        while (contador < n_filas)
                        {
                            Dgv_ASTM95_Run_Differentials.Rows.Add();
                            Dgv_ASTM_95_Differential.Rows.Add();
                            contador++;
                        }
                        //max%
                        int n_filas1 = Convert.ToInt32(Dgv_ASTM_Single_Aperture.Rows.Count.ToString()) + 1;
                        int contador1 = 1;
                        while (contador1 < n_filas1)
                        {
                            Dgv_Single_Aperture_Run_Differential.Rows.Add();
                            Dgv_Single_Aperture_Differential.Rows.Add();
                            contador1++;
                        }

                        //llenar la primera y la ultima columna 95%
                        double comp;
                        double val = 100000;
                        foreach (DataGridViewRow max in Dgv_ASTM_D95.Rows)
                        {
                            string max1 = Convert.ToString(max.Cells[2].Value);
                            comp = Convert.ToDouble(max1);
                            if (comp < val)
                            {
                                val = comp;
                                name = max.Cells[0].Value.ToString();
                                Dgv_ASTM95_Run_Differentials.Rows[0].Cells[0].Value = name;
                                Dgv_ASTM_95_Differential.Rows[0].Cells[0].Value = name;
                            }
                        }
                        Dgv_ASTM95_Run_Differentials.Rows[0].Cells[1].Value = Math.Round(val, 2);
                        Dgv_ASTM_95_Differential.Rows[0].Cells[1].Value = Math.Round(val, 2);

                        double comp1;
                        double val1 = 100000;
                        foreach (DataGridViewRow max1 in Dgv_ASTM_D95.Rows)
                        {
                            string max11 = Convert.ToString(max1.Cells[4].Value);
                            comp1 = Convert.ToDouble(max11);
                            if (comp1 < val1)
                            {
                                val1 = comp1;
                            }
                        }
                        Dgv_ASTM95_Run_Differentials.Rows[n_filas - 1].Cells[1].Value = Math.Round(val1, 2);
                        Dgv_ASTM_95_Differential.Rows[n_filas - 1].Cells[1].Value = Math.Round(val1, 2);
                        //llenar la primera y la ultima columna max%
                        double compz;
                        double valz = 100000;
                        foreach (DataGridViewRow max in Dgv_ASTM_Single_Aperture.Rows)
                        {
                            string max1 = Convert.ToString(max.Cells[2].Value);
                            compz = Convert.ToDouble(max1);
                            if (compz < valz)
                            {
                                valz = compz;
                                namez = max.Cells[0].Value.ToString();
                                Dgv_Single_Aperture_Run_Differential.Rows[0].Cells[0].Value = namez;
                                Dgv_Single_Aperture_Differential.Rows[0].Cells[0].Value = namez;
                            }
                        }
                        Dgv_Single_Aperture_Run_Differential.Rows[0].Cells[1].Value = Math.Round(valz, 2);
                        Dgv_Single_Aperture_Differential.Rows[0].Cells[1].Value = Math.Round(valz, 2);

                        double comp1z;
                        double val1z = 100000;
                        foreach (DataGridViewRow max1 in Dgv_ASTM_Single_Aperture.Rows)
                        {
                            string max11 = Convert.ToString(max1.Cells[4].Value);
                            comp1z = Convert.ToDouble(max11);
                            if (comp1z < val1z)
                            {
                                val1z = comp1z;
                            }
                        }
                        Dgv_Single_Aperture_Run_Differential.Rows[n_filas1 - 1].Cells[1].Value = Math.Round(val1z, 2);
                        Dgv_Single_Aperture_Differential.Rows[n_filas1 - 1].Cells[1].Value = Math.Round(val1z, 2);

                        //segunda carrera 95%
                        double comp2;
                        double val2 = 100000;
                        foreach (DataGridViewRow max21 in Dgv_ASTM_D95.Rows)
                        {
                            string max2 = Convert.ToString(max21.Cells[3].Value);
                            comp2 = Convert.ToDouble(max2);
                            if (comp2 < val2)
                            {
                                val2 = comp2;
                            }
                        }
                        Dgv_ASTM95_Run_Differentials.Rows[0].Cells[2].Value = Math.Round(val2, 2);
                        Dgv_ASTM_95_Differential.Rows[0].Cells[2].Value = Math.Round(val2, 2);

                        double comp4;
                        double val4 = 100000;
                        foreach (DataGridViewRow max4 in Dgv_ASTM_D95.Rows)
                        {
                            string max41 = Convert.ToString(max4.Cells[5].Value);
                            comp4 = Convert.ToDouble(max41);
                            if (comp4 < val4)
                            {
                                val4 = comp4;
                                lname = max4.Cells[0].Value.ToString();
                                Dgv_ASTM95_Run_Differentials.Rows[n_filas - 1].Cells[0].Value = lname;
                                Dgv_ASTM_95_Differential.Rows[n_filas - 1].Cells[0].Value = lname;
                            }
                        }
                        Dgv_ASTM95_Run_Differentials.Rows[n_filas - 1].Cells[2].Value = Math.Round(val4, 2);
                        Dgv_ASTM_95_Differential.Rows[n_filas - 1].Cells[2].Value = Math.Round(val4, 2);
                        //segunda carrera max%
                        double comp2z;
                        double val2z = 100000;
                        foreach (DataGridViewRow max21 in Dgv_ASTM_Single_Aperture.Rows)
                        {
                            string max2 = Convert.ToString(max21.Cells[3].Value);
                            comp2z = Convert.ToDouble(max2);
                            if (comp2z < val2z)
                            {
                                val2z = comp2z;
                            }
                        }
                        Dgv_Single_Aperture_Run_Differential.Rows[0].Cells[2].Value = Math.Round(val2z, 2);
                        Dgv_Single_Aperture_Differential.Rows[0].Cells[2].Value = Math.Round(val2z, 2);

                        double comp4z;
                        double val4z = 100000;
                        foreach (DataGridViewRow max4 in Dgv_ASTM_Single_Aperture.Rows)
                        {
                            string max41 = Convert.ToString(max4.Cells[5].Value);
                            comp4z = Convert.ToDouble(max41);
                            if (comp4z < val4z)
                            {
                                val4z = comp4z;
                                lnamez = max4.Cells[0].Value.ToString();
                                Dgv_Single_Aperture_Run_Differential.Rows[n_filas1 - 1].Cells[0].Value = lnamez;
                                Dgv_Single_Aperture_Differential.Rows[n_filas1 - 1].Cells[0].Value = lnamez;
                            }
                        }
                        Dgv_Single_Aperture_Run_Differential.Rows[n_filas1 - 1].Cells[2].Value = Math.Round(val4z, 2);
                        Dgv_Single_Aperture_Differential.Rows[n_filas1 - 1].Cells[2].Value = Math.Round(val4z, 2);

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
                        foreach (DataGridViewRow con in Dgv_ASTM_D95.Rows)
                        {
                            //Primera Corrida
                            while (com < 1)
                            {
                                acumulador = Convert.ToDouble(Dgv_ASTM95_Run_Differentials.Rows[0].Cells[1].Value);
                                com++;
                            }
                            if (con.Cells[0].Value.ToString() != name.ToString())
                            {
                                //Operaciones
                                calculo = ((Convert.ToDouble(con.Cells[2].Value)) - (Convert.ToDouble(Dgv_ASTM_D95.Rows[check1].Cells[2].Value.ToString())));
                                Dgv_ASTM95_Run_Differentials.Rows[check].Cells[1].Value = Math.Round(calculo, 2);
                                Dgv_ASTM_95_Differential.Rows[check].Cells[1].Value = Math.Round(calculo, 2);

                                //Aumento de acumulador
                                acumulador = acumulador + Convert.ToDouble(con.Cells[2].Value.ToString());
                            }

                            //Segunda Corrida
                            while (com1 < 1)
                            {
                                acumulador1 = Convert.ToDouble(Dgv_ASTM95_Run_Differentials.Rows[0].Cells[2].Value);
                                com1++;
                            }
                            if (con.Cells[0].Value.ToString() != name.ToString())
                            {
                                //Operaciones
                                calculo1 = ((Convert.ToDouble(con.Cells[3].Value)) - (Convert.ToDouble(Dgv_ASTM_D95.Rows[check1].Cells[3].Value.ToString())));
                                Dgv_ASTM95_Run_Differentials.Rows[check].Cells[2].Value = Math.Round(calculo1, 2);
                                Dgv_ASTM_95_Differential.Rows[check].Cells[2].Value = Math.Round(calculo1, 2);

                                //Aumento de acumulador
                                acumulador1 = acumulador1 + Convert.ToDouble(con.Cells[3].Value.ToString());
                                check++;
                                check1++;
                            }
                        }

                        //Llenar los ultimos datos max%
                        foreach (DataGridViewRow con in Dgv_ASTM_Single_Aperture.Rows)
                        {
                            //Primera Corrida
                            while (comz < 1)
                            {
                                acumuladorz = Convert.ToDouble(Dgv_Single_Aperture_Run_Differential.Rows[0].Cells[1].Value);
                                comz++;
                            }
                            if (con.Cells[0].Value.ToString() != namez.ToString())
                            {
                                //Operaciones
                                calculoz = ((Convert.ToDouble(con.Cells[2].Value)) - (Convert.ToDouble(Dgv_ASTM_Single_Aperture.Rows[check1z].Cells[2].Value.ToString())));
                                Dgv_Single_Aperture_Run_Differential.Rows[checkz].Cells[1].Value = Math.Round(calculoz, 2);
                                Dgv_Single_Aperture_Differential.Rows[checkz].Cells[1].Value = Math.Round(calculoz, 2);

                                //Aumento de acumulador
                                acumuladorz = acumuladorz + Convert.ToDouble(con.Cells[2].Value.ToString());
                            }

                            //Segunda Corrida
                            while (com1z < 1)
                            {
                                acumulador1z = Convert.ToDouble(Dgv_Single_Aperture_Run_Differential.Rows[0].Cells[2].Value);
                                com1z++;
                            }
                            if (con.Cells[0].Value.ToString() != namez.ToString())
                            {
                                //Operaciones
                                calculo1z = ((Convert.ToDouble(con.Cells[3].Value)) - (Convert.ToDouble(Dgv_ASTM_Single_Aperture.Rows[check1z].Cells[3].Value.ToString())));
                                Dgv_Single_Aperture_Run_Differential.Rows[checkz].Cells[2].Value = Math.Round(calculo1z, 2);
                                Dgv_Single_Aperture_Differential.Rows[checkz].Cells[2].Value = Math.Round(calculo1z, 2);

                                //Aumento de acumulador
                                acumulador1z = acumulador1z + Convert.ToDouble(con.Cells[3].Value.ToString());
                                checkz++;
                                check1z++;
                            }
                        }
                        renombrar1();
                    }
                    Dgv_ASTM95_Run_Differentials.Columns[3].Visible = false;
                    Dgv_Single_Aperture_Run_Differential.Columns[3].Visible = false;
                }
                else if (Dgv_Particle_Data.Rows[0].Cells[2].Value.ToString() == "Run_1 (Vol%)")
                {
                    //1 corrida
                    if (Dgv_ASTM_D95.Rows.Count == 2)
                    {
                        Dgv_ASTM95_Run_Differentials.Visible = true;
                        Dgv_Single_Aperture_Run_Differential.Visible = true;

                        Manage_Data data = new Manage_Data();
                        data.addRowsToDataGridView(Dgv_ASTM95_Run_Differentials);
                        data.addRowsToDataGridView(Dgv_ASTM_95_Differential);

                        data.addRowsToDataGridView(Dgv_Single_Aperture_Run_Differential);
                        data.addRowsToDataGridView(Dgv_Single_Aperture_Differential);

                        Differential differential = new Differential();
                        differential.assignComparisonVariable(Dgv_ASTM_D95, Dgv_ASTM95_Run_Differentials);
                        differential.assignComparisonVariable(Dgv_ASTM_D95, Dgv_ASTM_95_Differential);

                        differential.assignComparisonVariable(Dgv_ASTM_Single_Aperture, Dgv_Single_Aperture_Run_Differential);
                        differential.assignComparisonVariable(Dgv_ASTM_Single_Aperture, Dgv_Single_Aperture_Differential);

                        differential.assignComparisonVariableRunTwo(Dgv_ASTM_D95, Dgv_ASTM95_Run_Differentials);
                        differential.assignComparisonVariableRunTwo(Dgv_ASTM_D95, Dgv_ASTM_95_Differential);

                        differential.assignComparisonVariableRunTwo(Dgv_ASTM_Single_Aperture, Dgv_Single_Aperture_Run_Differential);
                        differential.assignComparisonVariableRunTwo(Dgv_ASTM_Single_Aperture, Dgv_Single_Aperture_Differential);

                        differential.createDifferential(Dgv_ASTM95_Run_Differentials);
                        differential.createDifferential(Dgv_ASTM_95_Differential);

                        differential.createDifferential(Dgv_Single_Aperture_Run_Differential);
                        differential.createDifferential(Dgv_Single_Aperture_Differential);

                        renombrar();
                    }
                    else if (Dgv_ASTM_D95.Rows.Count > 2)
                    {
                        Dgv_ASTM95_Run_Differentials.Visible = true;
                        Dgv_Single_Aperture_Run_Differential.Visible = true;

                        //95%
                        int n_filas = Convert.ToInt32(Dgv_ASTM_D95.Rows.Count.ToString()) + 1;
                        int contador = 1;
                        while (contador < n_filas)
                        {
                            Dgv_ASTM95_Run_Differentials.Rows.Add();
                            Dgv_ASTM_95_Differential.Rows.Add();
                            contador++;
                        }
                        //max%
                        int n_filas1 = Convert.ToInt32(Dgv_ASTM_Single_Aperture.Rows.Count.ToString()) + 1;
                        int contador1 = 1;
                        while (contador1 < n_filas1)
                        {
                            Dgv_Single_Aperture_Run_Differential.Rows.Add();
                            Dgv_Single_Aperture_Differential.Rows.Add();
                            contador1++;
                        }

                        //llenar la primera y la ultima columna 95%
                        double comp;
                        double val = 100000;
                        foreach (DataGridViewRow max in Dgv_ASTM_D95.Rows)
                        {
                            string max1 = Convert.ToString(max.Cells[2].Value);
                            comp = Convert.ToDouble(max1);
                            if (comp < val)
                            {
                                val = comp;
                                name = max.Cells[0].Value.ToString();
                                Dgv_ASTM95_Run_Differentials.Rows[0].Cells[0].Value = name;
                                Dgv_ASTM_95_Differential.Rows[0].Cells[0].Value = name;
                            }
                        }
                        Dgv_ASTM95_Run_Differentials.Rows[0].Cells[1].Value = Math.Round(val, 2);
                        Dgv_ASTM_95_Differential.Rows[0].Cells[1].Value = Math.Round(val, 2);
                        //llenar la primera y la ultima columna max%
                        double comp1;
                        double val1 = 100000;
                        foreach (DataGridViewRow max in Dgv_ASTM_Single_Aperture.Rows)
                        {
                            string max1 = Convert.ToString(max.Cells[2].Value);
                            comp1 = Convert.ToDouble(max1);
                            if (comp1 < val1)
                            {
                                val1 = comp1;
                                namez = max.Cells[0].Value.ToString();
                                Dgv_Single_Aperture_Run_Differential.Rows[0].Cells[0].Value = namez;
                                Dgv_Single_Aperture_Differential.Rows[0].Cells[0].Value = namez;
                            }
                        }
                        Dgv_Single_Aperture_Run_Differential.Rows[0].Cells[1].Value = Math.Round(val1, 2);
                        Dgv_Single_Aperture_Differential.Rows[0].Cells[1].Value = Math.Round(val1, 2);

                        //2 95%
                        double comp1z;
                        double val1z = 100000;
                        foreach (DataGridViewRow max1 in Dgv_ASTM_D95.Rows)
                        {
                            string max11 = Convert.ToString(max1.Cells[3].Value);
                            comp1z = Convert.ToDouble(max11);
                            if (comp1z < val1z)
                            {
                                val1z = comp1z;
                            }
                        }
                        Dgv_ASTM95_Run_Differentials.Rows[n_filas - 1].Cells[1].Value = Math.Round(val1z, 2);
                        Dgv_ASTM_95_Differential.Rows[n_filas - 1].Cells[1].Value = Math.Round(val1z, 2);
                        //2 max%
                        double comp1z1;
                        double val1z1 = 100000;
                        foreach (DataGridViewRow max1 in Dgv_ASTM_Single_Aperture.Rows)
                        {
                            string max11 = Convert.ToString(max1.Cells[3].Value);
                            comp1z1 = Convert.ToDouble(max11);
                            if (comp1z1 < val1z1)
                            {
                                val1z1 = comp1z1;
                            }
                        }
                        Dgv_Single_Aperture_Run_Differential.Rows[n_filas1 - 1].Cells[1].Value = Math.Round(val1z1, 2);
                        Dgv_Single_Aperture_Differential.Rows[n_filas1 - 1].Cells[1].Value = Math.Round(val1z1, 2);

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
                        foreach (DataGridViewRow con in Dgv_ASTM_D95.Rows)
                        {
                            //Primera Corrida
                            while (com < 1)
                            {
                                acumulador = Convert.ToDouble(Dgv_ASTM95_Run_Differentials.Rows[0].Cells[1].Value);
                                com++;
                            }
                            if (con.Cells[0].Value.ToString() != name.ToString())
                            {
                                //Operaciones
                                calculo = ((Convert.ToDouble(con.Cells[2].Value)) - (Convert.ToDouble(Dgv_ASTM_D95.Rows[check1].Cells[2].Value.ToString())));
                                Dgv_ASTM95_Run_Differentials.Rows[check].Cells[1].Value = Math.Round(calculo, 2);
                                Dgv_ASTM_95_Differential.Rows[check].Cells[1].Value = Math.Round(calculo, 2);

                                //Aumento de acumulador
                                acumulador = acumulador + Convert.ToDouble(con.Cells[2].Value.ToString());
                                check++;
                                check1++;
                            }
                        }
                        //Llenar los ultimos datos max%
                        foreach (DataGridViewRow con in Dgv_ASTM_Single_Aperture.Rows)
                        {
                            //Primera Corrida
                            while (comz < 1)
                            {
                                acumuladorz = Convert.ToDouble(Dgv_Single_Aperture_Run_Differential.Rows[0].Cells[1].Value);
                                comz++;
                            }
                            if (con.Cells[0].Value.ToString() != namez.ToString())
                            {
                                //Operaciones
                                calculoz = ((Convert.ToDouble(con.Cells[2].Value)) - (Convert.ToDouble(Dgv_ASTM_Single_Aperture.Rows[check1z].Cells[2].Value.ToString())));
                                Dgv_Single_Aperture_Run_Differential.Rows[checkz].Cells[1].Value = Math.Round(calculoz, 2);
                                Dgv_Single_Aperture_Differential.Rows[checkz].Cells[1].Value = Math.Round(calculoz, 2);

                                //Aumento de acumulador
                                acumuladorz = acumuladorz + Convert.ToDouble(con.Cells[2].Value.ToString());
                                checkz++;
                                check1z++;
                            }
                        }
                        renombrar1();
                    }
                    Dgv_ASTM95_Run_Differentials.Columns[3].Visible = false;
                    Dgv_ASTM95_Run_Differentials.Columns[2].Visible = false;
                    Dgv_Single_Aperture_Run_Differential.Columns[3].Visible = false;
                    Dgv_Single_Aperture_Run_Differential.Columns[2].Visible = false;
                }
            }
            catch (Exception tr)
            {
                Dgv_ASTM95_Run_Differentials.Visible = false;
                Dgv_Single_Aperture_Run_Differential.Visible = false;
                MessageBox.Show("It's necessary to mark the cumulative");
            }

            Dgv_ASTM_D95_Accumulated_rigth_left.Visible = true;
            Dgv_ASTM_95_Differential.Visible = true;
            Btn_Hide_Cumulatives.Visible = true;
            allowSelect = true;
            TabControl_Main_Menu.SelectedTab = Page_Report_View;
            allowSelect = false;
            Dgv_ASTM_D95_Accumulated_rigth_left.AllowUserToAddRows = false;

            Dgv_Single_Aperture_Accumulated_right_left.Visible = true;
            Dgv_Single_Aperture_Differential.Visible = true;
            Dgv_Single_Aperture_Accumulated_right_left.AllowUserToAddRows = false;

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
            foreach (DataGridViewRow row in Dgv_Sample_Information.Rows)
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
            if (Dgv_Particle_Data.Rows[0].Cells[4].Value.ToString() == "Run_3 (Vol%)")
            {
                //3 corridas
                //Asignacion a variables
                if (Nombres.Count > 0)
                {
                    n1 = Nombres[0];
                    n20 = Nombres[1];
                    n30 = Nombres[2];

                    Lbl_Name_Value_Run_One.Text = n1;
                    label31.Text = n20;
                    label49.Text = n30;
                }
                else
                {
                    Lbl_Name_Value_Run_One.Text = "";
                    label31.Text = "";
                    label49.Text = "";

                    Lbl_Name_Run_One.Visible = false;
                    label40.Visible = false;
                    label58.Visible = false;
                }
                if (Fecha.Count > 0)
                {
                    f1 = Fecha[0];
                    f2 = Fecha[1];
                    f3 = Fecha[2];

                    Lbl_Sample_Data_Value_Run_One.Text = f1;
                    label30.Text = f2;
                    label48.Text = f3;
                }
                else
                {
                    Lbl_Sample_Data_Value_Run_One.Text = "";
                    label30.Text = "";
                    label48.Text = "";
                }
                if (Usuarios.Count > 0)
                {
                    u1 = Usuarios[0];
                    u2 = Usuarios[1];
                    u3 = Usuarios[2];

                    Lbl_User_Value_Run_One.Text = u1;
                    label29.Text = u2;
                    label47.Text = u3;
                }
                else
                {
                    Lbl_User_Value_Run_One.Text = "";
                    label29.Text = "";
                    label47.Text = "";
                }
                if (Equipos.Count > 0)
                {
                    e1 = Equipos[0];
                    e2 = Equipos[1];
                    e3 = Equipos[2];

                    Lbl_Device_Value_Run_One.Text = e1;
                    label28.Text = e2;
                    label46.Text = e3;
                }
                else
                {
                    Lbl_Device_Value_Run_One.Text = "";
                    label28.Text = "";
                    label46.Text = "";
                }
                if (Ids.Count > 0)
                {
                    i1 = Ids[0];
                    i2 = Ids[1];
                    i3 = Ids[2];

                    Lbl_Sample_Id_Value_Run_One.Text = i1;
                    label27.Text = i2;
                    label45.Text = i3;
                }
                else
                {
                    Lbl_Sample_Id_Value_Run_One.Text = "";
                    label27.Text = "";
                    label45.Text = "";
                }
                if (Grupos.Count > 0)
                {
                    g1 = Grupos[0];
                    g2 = Grupos[1];
                    g3 = Grupos[2];

                    Lbl_Group_Id_Value.Text = g1;
                    label26.Text = g2;
                    label44.Text = g3;
                }
                else
                {
                    Lbl_Group_Id_Value.Text = "";
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
            else if (Dgv_Particle_Data.Rows[0].Cells[3].Value.ToString() == "Run_2 (Vol%)")
            {
                // 2 corridas
                //Asignacion a variables
                if (Nombres.Count > 0)
                {
                    n1 = Nombres[0];
                    n20 = Nombres[1];

                    Lbl_Name_Value_Run_One.Text = n1;
                    label31.Text = n20;
                    label49.Text = "";
                    label60.Text = "";
                    label58.Text = "";
                }
                else
                {
                    Lbl_Name_Value_Run_One.Text = "";
                    label31.Text = "";
                    label49.Text = "";

                    Lbl_Name_Run_One.Visible = false;
                    label40.Visible = false;
                    label58.Visible = false;
                }
                if (Fecha.Count > 0)
                {
                    f1 = Fecha[0];
                    f2 = Fecha[1];

                    Lbl_Sample_Data_Value_Run_One.Text = f1;
                    label30.Text = f2;
                    label48.Text = "";
                    label57.Text = "";
                }
                else
                {
                    Lbl_Sample_Data_Value_Run_One.Text = "";
                    label30.Text = "";
                    label48.Text = "";
                }
                if (Usuarios.Count > 0)
                {
                    u1 = Usuarios[0];
                    u2 = Usuarios[1];

                    Lbl_User_Value_Run_One.Text = u1;
                    label29.Text = u2;
                    label47.Text = "";
                    label56.Text = "";
                }
                else
                {
                    Lbl_User_Value_Run_One.Text = "";
                    label29.Text = "";
                    label47.Text = "";
                }
                if (Equipos.Count > 0)
                {
                    e1 = Equipos[0];
                    e2 = Equipos[1];

                    Lbl_Device_Value_Run_One.Text = e1;
                    label28.Text = e2;
                    label46.Text = "";
                    label55.Text = "";
                }
                else
                {
                    Lbl_Device_Value_Run_One.Text = "";
                    label28.Text = "";
                    label46.Text = "";
                }
                if (Ids.Count > 0)
                {
                    i1 = Ids[0];
                    i2 = Ids[1];

                    Lbl_Sample_Id_Value_Run_One.Text = i1;
                    label27.Text = i2;
                    label45.Text = "";
                    label54.Text = "";
                }
                else
                {
                    Lbl_Sample_Id_Value_Run_One.Text = "";
                    label27.Text = "";
                    label45.Text = "";
                }
                if (Grupos.Count > 0)
                {
                    g1 = Grupos[0];
                    g2 = Grupos[1];

                    Lbl_Group_Id_Value.Text = g1;
                    label26.Text = g2;
                    label44.Text = "";
                    label53.Text = "";
                }
                else
                {
                    Lbl_Group_Id_Value.Text = "";
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
            else if (Dgv_Particle_Data.Rows[0].Cells[2].Value.ToString() == "Run_1 (Vol%)")
            {
                //1 corrida
                //Asignacion a variables
                if (Nombres.Count > 0)
                {
                    n1 = Nombres[0];

                    Lbl_Name_Value_Run_One.Text = n1;
                    label31.Text = "";
                    label49.Text = "";
                    label40.Text = "";
                    label58.Text = "";
                }
                else
                {
                    Lbl_Name_Value_Run_One.Text = "";
                    label31.Text = "";
                    label49.Text = "";

                    Lbl_Name_Run_One.Visible = false;
                    label40.Visible = false;
                    label58.Visible = false;
                }
                if (Fecha.Count > 0)
                {
                    f1 = Fecha[0];

                    Lbl_Sample_Data_Value_Run_One.Text = f1;
                    label30.Text = "";
                    label48.Text = "";
                    label39.Text = "";
                    label57.Text = "";
                }
                else
                {
                    Lbl_Sample_Data_Value_Run_One.Text = "";
                    label30.Text = "";
                    label48.Text = "";
                }
                if (Usuarios.Count > 0)
                {
                    u1 = Usuarios[0];

                    Lbl_User_Value_Run_One.Text = u1;
                    label29.Text = "";
                    label47.Text = "";
                    label38.Text = "";
                    label56.Text = "";
                }
                else
                {
                    Lbl_User_Value_Run_One.Text = "";
                    label29.Text = "";
                    label47.Text = "";
                }
                if (Equipos.Count > 0)
                {
                    e1 = Equipos[0];

                    Lbl_Device_Value_Run_One.Text = e1;
                    label28.Text = "";
                    label46.Text = "";
                    label37.Text = "";
                    label55.Text = "";
                }
                else
                {
                    Lbl_Device_Value_Run_One.Text = "";
                    label28.Text = "";
                    label46.Text = "";
                }
                if (Ids.Count > 0)
                {
                    i1 = Ids[0];

                    Lbl_Sample_Id_Value_Run_One.Text = i1;
                    label27.Text = "";
                    label45.Text = "";
                    label36.Text = "";
                    label54.Text = "";
                }
                else
                {
                    Lbl_Sample_Id_Value_Run_One.Text = "";
                    label27.Text = "";
                    label45.Text = "";
                }
                if (Grupos.Count > 0)
                {
                    g1 = Grupos[0];

                    Lbl_Group_Id_Value.Text = g1;
                    label26.Text = "";
                    label44.Text = "";
                    label35.Text = "";
                    label53.Text = "";
                }
                else
                {
                    Lbl_Group_Id_Value.Text = "";
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
            Dgv_ASTM_D95_Accumulated_rigth_left.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            Dgv_ASTM_D95_Accumulated_rigth_left.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;

            Dgv_ASTM_95_Differential.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            Dgv_ASTM_95_Differential.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;

            //Asignacion de Datos de la empresa max%
            Dgv_Single_Aperture_Accumulated_right_left.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            Dgv_Single_Aperture_Accumulated_right_left.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;

            Dgv_Single_Aperture_Differential.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            Dgv_Single_Aperture_Differential.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;

            //Insertar 999 en dgv 5
            Dgv_ASTM_D95_Accumulated_rigth_left.Rows.Insert(0, "", "999", "0", "0", "0", "0", "0", "0");
            Dgv_Single_Aperture_Accumulated_right_left.Rows.Insert(0, "", "999", "0", "0", "0", "0", "0", "0");

            Dgv_ASTM_D95_Accumulated_rigth_left.ReadOnly = true;
            Dgv_ASTM_95_Differential.ReadOnly = true;
            Dgv_ASTM_D95_Accumulated_rigth_left.ClearSelection();
            Dgv_ASTM_95_Differential.ClearSelection();

            Dgv_Single_Aperture_Accumulated_right_left.ReadOnly = true;
            Dgv_Single_Aperture_Differential.ReadOnly = true;
            Dgv_Single_Aperture_Accumulated_right_left.ClearSelection();
            Dgv_Single_Aperture_Differential.ClearSelection();
        }

        private void addColumnsOfCummulativeValues(DataGridView dataGridView, string numberOfRun)
        {
            Manage_Data manageData = new Manage_Data();

            if (numberOfRun.Equals("3"))
            {
                manageData.addColumnToDatagridView("Run_1 Cumulative <", dataGridView);
                manageData.addColumnToDatagridView("Run_2 Cumulative <", dataGridView);
                manageData.addColumnToDatagridView("Run_3 Cumulative <", dataGridView);
            }
            if (numberOfRun.Equals("2"))
            {
                manageData.addColumnToDatagridView("Run_1 Cumulative <", dataGridView);
                manageData.addColumnToDatagridView("Run_2 Cumulative <", dataGridView);
            }
            if (numberOfRun.Equals("1"))
            {
                manageData.addColumnToDatagridView("Run_1 Cumulative <", dataGridView);
            }
            
        }

        private void addColumnsOfCummulativeValuesToLeft(DataGridView dataGridView, string numberOfRun)
        {
            Manage_Data manageData = new Manage_Data();
            if (numberOfRun.Equals("3"))
            {
                manageData.addColumnToDatagridView("Run_1 Cumulative >", dataGridView);
                manageData.addColumnToDatagridView("Run_2 Cumulative >", dataGridView);
                manageData.addColumnToDatagridView("Run_3 Cumulative >", dataGridView);
            }
            if (numberOfRun.Equals("2"))
            {
                manageData.addColumnToDatagridView("Run_1 Cumulative >", dataGridView);
                manageData.addColumnToDatagridView("Run_2 Cumulative >", dataGridView);
            }
            if (numberOfRun.Equals("1"))
            {
                manageData.addColumnToDatagridView("Run_1 Cumulative >", dataGridView);
            }          
        }


        private void addCumulativeValuesForEachRunToDataGridView()
        {
            for (int numberOfRun = 2; numberOfRun<= 4; numberOfRun++)
            {
                this.addCumulativeValuesToRightOfDataGridView(Dgv_ASTM95_Detector_Number, Dgv_ASTM_D95, numberOfRun);
                this.addCumulativeValuesToRightOfDataGridView(Dgv_ASTM95_Detector_Number, Dgv_ASTM_D95_Accumulated_rigth_left, numberOfRun);

                this.addCumulativeValuesToRightOfDataGridView(Dgv_Single_Aperture_Detector, Dgv_ASTM_Single_Aperture, numberOfRun);
                this.addCumulativeValuesToRightOfDataGridView(Dgv_Single_Aperture_Detector, Dgv_Single_Aperture_Accumulated_right_left, numberOfRun);
            }       
        }

        private void addCumulativeValuesToRightOfDataGridView(DataGridView dgvToReview, DataGridView dgvToAddValues, int numberOfRun)
        {
            foreach (DataGridViewRow row in dgvToReview.Rows)
            {
                double accumulated = 0;
                int n = 1;
                //aumentar a la fila los valores acumulativos a la derecha (los que van arriba)
                try
                {
                    int detectorNumberColumnValue = Convert.ToInt32(row.Cells[3].Value);
                    while (n <= detectorNumberColumnValue)
                    {
                        double valueInColumnRunOne = Convert.ToDouble(Dgv_Particle_Data.Rows[n].Cells[numberOfRun].Value);
                        accumulated = accumulated + valueInColumnRunOne;
                        n++;
                        if (accumulated > 100)
                        {
                            accumulated = 100;
                        }
                        dgvToAddValues.Rows[row.Index].Cells[numberOfRun].Value = Math.Round(accumulated, 2);
                    }

                    //Asignacion de valores para interpolación
                    this.assignmentOfValuesForInterpolation(n, row, accumulated, dgvToAddValues, numberOfRun);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void assignmentOfValuesForInterpolation(int numberOfRow, DataGridViewRow row, double accumulated, DataGridView dgvToAddValues, int numberofRun)
        {
            //Asignacion de valores para interpolación
            Micron _micron = new Micron();

            double micronInferiorLimitValue = Convert.ToDouble(Dgv_Particle_Data.Rows[numberOfRow - 1].Cells[0].Value.ToString());
            double micronValue = Convert.ToDouble(Dgv_Particle_Data.Rows[numberOfRow].Cells[0].Value.ToString());

            double runOneValue = Convert.ToDouble(Dgv_Particle_Data.Rows[numberOfRow].Cells[numberofRun].Value);
            double totalAccumulated = accumulated + runOneValue;
            totalAccumulated = Convert.ToDouble(Math.Round(totalAccumulated, 2).ToString());

            string micronValueInString = Convert.ToString(row.Cells[2].Value);
            double micron = _micron.getRoundedMicron(micronValueInString);
            double result = this.interpolationFormula(micron, micronValue, micronInferiorLimitValue, accumulated, totalAccumulated); //74.7019 -- 29- 31.50, 28.70 74.41  77.15

            dgvToAddValues.Rows[row.Index].Cells[numberofRun].Value = Math.Round(result, 2);
        }

        private double interpolationFormula(double micron, double micronValue, double micronInferiorLimitValue, double accumulated, double roundedValue)
        {
            double arriba = micron - micronInferiorLimitValue;
            double abajo = micronValue - micronInferiorLimitValue;
            double division = arriba / abajo;
            return  accumulated + (division * (roundedValue - accumulated));
        }

        public void renombrar()
        {
            //Renombrar 95%
            //Se renombran las celdas 0 de los grids 4, 6
            foreach (DataGridViewRow row in Dgv_ASTM95_Run_Differentials.Rows)
            {
                try
                {
                    row.Cells[0].Value = Dgv_ASTM_D95.Rows[row.Index].Cells[1].Value;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            foreach (DataGridViewRow renombre1 in Dgv_ASTM_95_Differential.Rows)
            {
                try
                {
                    renombre1.Cells[0].Value = Dgv_ASTM_D95.Rows[renombre1.Index].Cells[1].Value;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            //Renombrar max%
            //Se renombran las celdas 0 de los grids 12, 15
            foreach (DataGridViewRow renombre in Dgv_Single_Aperture_Run_Differential.Rows)
            {
                try
                {
                    renombre.Cells[0].Value = Dgv_ASTM_Single_Aperture.Rows[renombre.Index].Cells[1].Value;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            foreach (DataGridViewRow renombre1 in Dgv_Single_Aperture_Differential.Rows)
            {
                try
                {
                    renombre1.Cells[0].Value = Dgv_ASTM_Single_Aperture.Rows[renombre1.Index].Cells[1].Value;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            try
            {
                //Para 95%
                Dgv_ASTM95_Run_Differentials.Rows[(Dgv_ASTM95_Run_Differentials.Rows.Count - 4)].Cells[0].Value = ("999");
                Dgv_ASTM95_Run_Differentials.Rows[(Dgv_ASTM95_Run_Differentials.Rows.Count - 3)].Cells[0].Value = (Dgv_ASTM_D95.Rows[(Dgv_ASTM_D95.Rows.Count - 2)].Cells[1].Value);
                Dgv_ASTM95_Run_Differentials.Rows[(Dgv_ASTM95_Run_Differentials.Rows.Count - 2)].Cells[0].Value = (Dgv_ASTM_D95.Rows[(Dgv_ASTM_D95.Rows.Count - 1)].Cells[1].Value);

                Dgv_ASTM_95_Differential.Rows[(Dgv_ASTM95_Run_Differentials.Rows.Count - 4)].Cells[0].Value = ("999");
                Dgv_ASTM_95_Differential.Rows[(Dgv_ASTM95_Run_Differentials.Rows.Count - 3)].Cells[0].Value = (Dgv_ASTM_D95.Rows[(Dgv_ASTM_D95.Rows.Count - 2)].Cells[1].Value);
                Dgv_ASTM_95_Differential.Rows[(Dgv_ASTM95_Run_Differentials.Rows.Count - 2)].Cells[0].Value = (Dgv_ASTM_D95.Rows[(Dgv_ASTM_D95.Rows.Count - 1)].Cells[1].Value);

                //Para max%
                Dgv_Single_Aperture_Run_Differential.Rows[(Dgv_Single_Aperture_Run_Differential.Rows.Count - 4)].Cells[0].Value = ("999");
                Dgv_Single_Aperture_Run_Differential.Rows[(Dgv_Single_Aperture_Run_Differential.Rows.Count - 3)].Cells[0].Value = (Dgv_ASTM_Single_Aperture.Rows[(Dgv_ASTM_Single_Aperture.Rows.Count - 2)].Cells[1].Value);
                Dgv_Single_Aperture_Run_Differential.Rows[(Dgv_Single_Aperture_Run_Differential.Rows.Count - 2)].Cells[0].Value = (Dgv_ASTM_Single_Aperture.Rows[(Dgv_ASTM_Single_Aperture.Rows.Count - 1)].Cells[1].Value);

                Dgv_Single_Aperture_Differential.Rows[(Dgv_Single_Aperture_Run_Differential.Rows.Count - 4)].Cells[0].Value = ("999");
                Dgv_Single_Aperture_Differential.Rows[(Dgv_Single_Aperture_Run_Differential.Rows.Count - 3)].Cells[0].Value = (Dgv_ASTM_Single_Aperture.Rows[(Dgv_ASTM_Single_Aperture.Rows.Count - 2)].Cells[1].Value);
                Dgv_Single_Aperture_Differential.Rows[(Dgv_Single_Aperture_Run_Differential.Rows.Count - 2)].Cells[0].Value = (Dgv_ASTM_Single_Aperture.Rows[(Dgv_ASTM_Single_Aperture.Rows.Count - 1)].Cells[1].Value);

            }
            catch (Exception ex)
            {
                MessageBox.Show("Mensaje " + ex.Message);
            }
        }

        public void renombrar1()
        {
            //Para 95%
            //Se renombran las celdas 0 de los grids 4, 6
            foreach (DataGridViewRow renombre in Dgv_ASTM95_Run_Differentials.Rows)
            {
                try
                {
                    renombre.Cells[0].Value = Dgv_ASTM_D95.Rows[renombre.Index].Cells[1].Value;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            foreach (DataGridViewRow renombre1 in Dgv_ASTM_95_Differential.Rows)
            {
                try
                {
                    renombre1.Cells[0].Value = Dgv_ASTM_D95.Rows[renombre1.Index].Cells[1].Value;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            //Para max%
            //Se renombran las celdas 0 de los grids 12, 5
            foreach (DataGridViewRow renombre in Dgv_Single_Aperture_Run_Differential.Rows)
            {
                try
                {
                    renombre.Cells[0].Value = Dgv_ASTM_Single_Aperture.Rows[renombre.Index].Cells[1].Value;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            foreach (DataGridViewRow renombre1 in Dgv_Single_Aperture_Differential.Rows)
            {
                try
                {
                    renombre1.Cells[0].Value = Dgv_ASTM_Single_Aperture.Rows[renombre1.Index].Cells[1].Value;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            //Recorrido de los nombres hacia abajo dependiendo del numero de filas que haya 
            //Para 95%
            try
            {
                int rep = Convert.ToInt32(Dgv_ASTM95_Run_Differentials.RowCount) + 2;
                int cont = 2;
                int cont2 = 1;
                while (cont < rep)
                {
                    Dgv_ASTM95_Run_Differentials.Rows[Dgv_ASTM95_Run_Differentials.RowCount - cont2].Cells[0].Value = Dgv_ASTM95_Run_Differentials.Rows[Dgv_ASTM95_Run_Differentials.RowCount - cont].Cells[0].Value;
                    Dgv_ASTM_95_Differential.Rows[Dgv_ASTM95_Run_Differentials.RowCount - cont2].Cells[0].Value = Dgv_ASTM95_Run_Differentials.Rows[Dgv_ASTM95_Run_Differentials.RowCount - cont].Cells[0].Value;

                    cont++;
                    cont2++;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            //Para max%
            try
            {
                int rep1 = Convert.ToInt32(Dgv_Single_Aperture_Run_Differential.RowCount) + 2;
                int cont1 = 2;
                int cont21 = 1;
                while (cont1 < rep1)
                {
                    Dgv_Single_Aperture_Run_Differential.Rows[Dgv_Single_Aperture_Run_Differential.RowCount - cont21].Cells[0].Value = Dgv_Single_Aperture_Run_Differential.Rows[Dgv_Single_Aperture_Run_Differential.RowCount - cont1].Cells[0].Value;
                    Dgv_Single_Aperture_Differential.Rows[Dgv_Single_Aperture_Run_Differential.RowCount - cont21].Cells[0].Value = Dgv_Single_Aperture_Run_Differential.Rows[Dgv_Single_Aperture_Run_Differential.RowCount - cont1].Cells[0].Value;

                    cont1++;
                    cont21++;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            //Valores por Default en las celdas 0 row 0
            Dgv_ASTM95_Run_Differentials.Rows[0].Cells[0].Value = "999";
            Dgv_ASTM_95_Differential.Rows[0].Cells[0].Value = "999";
            Dgv_Single_Aperture_Run_Differential.Rows[0].Cells[0].Value = "999";
            Dgv_Single_Aperture_Differential.Rows[0].Cells[0].Value = "999";
        }

        private void Hide_Differential_Click(object sender, EventArgs e)
        {
            diffferential = "si";
            try
            {
                foreach (DataGridViewRow row in Dgv_ASTM_95_Differential.Rows)
                {
                    foreach (DataGridViewColumn col in Dgv_ASTM_95_Differential.Columns)
                    {
                        Dgv_ASTM_95_Differential.Rows[row.Index].Cells[col.Index].Value = "";
                    }
                }

                foreach (DataGridViewRow row in Dgv_Single_Aperture_Differential.Rows)
                {
                    foreach (DataGridViewColumn col in Dgv_Single_Aperture_Differential.Columns)
                    {
                        Dgv_Single_Aperture_Differential.Rows[row.Index].Cells[col.Index].Value = "";
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            Btn_Hide_Differential.Visible = false;
            Dgv_ASTM_95_Differential.Visible = false;
            Dgv_Single_Aperture_Differential.Visible = false;
        }

        private void datos_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void Btn_Return_To_Excell_File_Click(object sender, EventArgs e)
        {
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
            TabControl_Main_Menu.SelectedTab = Page_Upload_Excell;
            allowSelect = false;
        }

        private void Btn_Clean_Data_Click(object sender, EventArgs e)
        {
            //Boton para borrar todos los registros y que se pueda volver a hacer una consulta con el mismo Excel pero con diferentes mallas seleccionadas
            Dgv_ASTM_D95.Visible = false;
            Dgv_ASTM95_Run_Differentials.Visible = false;

            Dgv_ASTM95_Detector_Number.Rows.Clear();
            Dgv_ASTM_D95.Rows.Clear();
            Dgv_ASTM95_Run_Differentials.Rows.Clear();
            Dgv_ASTM_D95_Accumulated_rigth_left.Rows.Clear();
            Dgv_ASTM_95_Differential.Rows.Clear();

            ch1 = true;
            ch2 = true;
            try
            {
                while (true)
                {
                    Dgv_ASTM_D95_Accumulated_rigth_left.Columns.RemoveAt(2);
                }
            }
            catch (Exception l)
            {

            }
            try
            {
                while (true)
                {
                    Dgv_ASTM_D95.Columns.RemoveAt(2);
                }
            }
            catch (Exception l)
            {

            }
            oc1 = true;
        }

    }
}
