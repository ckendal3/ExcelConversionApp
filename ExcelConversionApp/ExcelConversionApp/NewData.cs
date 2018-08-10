using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelConversionApp
{
    public class NewData
    {
        public NewData()
        {

        }

        // int respresents the cellId to place in and the value is the value of the newCell
        public Dictionary<int, string> stringData;
        public Dictionary<int, long> numericData;

    }
}
