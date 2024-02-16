using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LecturaExcel.Responsabilitis
{
    public class Micron
    {
        public double getRoundedMicron(string micron)
        {
            Match numbersInMicron = Regex.Match(micron, "(\\d+)");
            double roundedMicron = 0;
            if (numbersInMicron.Success)
            {
                roundedMicron = Convert.ToDouble(numbersInMicron.Value);
            }
            return roundedMicron;
        }

        public bool checkIfMicronHasIntegers(double micron)
        {
            if ((micron == 1) || (micron == 2) || (micron == 3) || (micron == 4))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

    }
}
