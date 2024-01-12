using ExcelDataReader;
using MaterialSkin;
using MaterialSkin.Controls;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LecturaExcel.View
{
    public partial class Load_File : MaterialForm
    {
        public Load_File()
        {
            InitializeComponent();
        }

        private void Load_File_Load(object sender, EventArgs e)
        {
            SkinManager.Theme = MaterialSkinManager.Themes.LIGHT;
            SkinManager.ColorScheme = new ColorScheme(Primary.Blue800, Primary.Blue700,
                Primary.Blue700, Accent.LightBlue100, TextShade.WHITE);
        }

        private void Btn_Load_File_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.ShowDialog();
            string path = openFileDialog.FileName.ToString();
            bool fileIsValid = checkIfFileIsValid(path);

            if (fileIsValid)
            {
                Form1 mainForm = new Form1(path);
                mainForm.Show();
            }
        }

        private bool checkIfFileIsValid(string filePath)
        {
            // ExcelPackage package = new ExcelPackage(new FileInfo(filePath));
            try
            {
                var stream = File.Open(filePath, FileMode.Open, FileAccess.Read);
                var reader = ExcelReaderFactory.CreateReader(stream);
                reader.Close();
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Elije un archivo valido "+ex.Message);
                return false;
            }
            
        }
    }
}
