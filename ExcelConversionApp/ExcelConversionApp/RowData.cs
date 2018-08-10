using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelConversionApp
{
    public class RowData
    {

        // int respresents the cellId to place in and the value is the value of the newCell
        public Dictionary<int, string> stringDict = new Dictionary<int, string>();
        public Dictionary<int, long> numericDict = new Dictionary<int, long>();

        public void AddString(int cell, string value)
        {
            if(stringDict.ContainsKey(cell))
            {
                Console.WriteLine("Cell data already exists");
                return;
            }

            stringDict.Add(cell, value);
        }

        public void AddNumber(int cell, int value)
        {
            if (numericDict.ContainsKey(cell))
            {
                Console.WriteLine("Cell data already exists");
                return;
            }

            numericDict.Add(cell, Convert.ToInt64(value));
        }

        public void AddNumber(int cell, long value)
        {
            if (numericDict.ContainsKey(cell))
            {
                Console.WriteLine("Cell data already exists");
                return;
            }

            numericDict.Add(cell, value);
        }

        public void AddNumber(int cell, double value)
        {
            if (numericDict.ContainsKey(cell))
            {
                Console.WriteLine("Cell data already exists");
                return;
            }

            numericDict.Add(cell, Convert.ToInt64(value));
        }
    }
}
