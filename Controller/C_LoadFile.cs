using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LecturaExcel.Controller
{
    public class C_LoadFile
    {
        public DataTable ParticleData;
        public DataTable SampleInformation;
        public void controll()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.ShowDialog();
            string filePath = openFileDialog.FileName.ToString();
            try
            {
                readExcelFile(filePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Selecciona un archivo excell con el formato correcto " +
                    "" + ex.Message.ToString());
            }
        }

        public void readExcelFile(string path)
        {
            var stream = File.Open(path, FileMode.Open, FileAccess.Read);
            var reader = ExcelReaderFactory.CreateReader(stream);
            var result = reader.AsDataSet();
            var tables = result.Tables.Cast<DataTable>();

            foreach (DataTable table in tables)
            {
                if (table.ToString() == "Data")
                {
                    ParticleData = table;
                }
                if (table.ToString() == "Sample Info")
                {
                    SampleInformation = table; 
                }
            }
        }

        public DataTable getParticleData()
        {
            return ParticleData;
        }
        public DataTable getSampleInformation()
        {
            return ParticleData;
        }
    }
}
