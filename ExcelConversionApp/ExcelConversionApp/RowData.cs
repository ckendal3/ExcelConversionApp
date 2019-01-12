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

    // ******************** NEEDS TO BE TESTED **************************
    public class RowDataV2
    {

        // int respresents the cellId to place in and the value is the value of the newCell
        public List<CellStruct> rowCellData = new List<CellStruct>();
        public List<CellCoords> rowCoords = new List<CellCoords>();


        /*
         * This needs to be tested along with a class (for holding coords and data) for possible simplification of code
         * public List<Tuple<CellCoords, CellStruct>> rowData = new List<Tuple<CellCoords, CellStruct>>();
        */

        /// <summary>
        /// Add a string value to a cell
        /// </summary>
        /// <param name="column"></param>
        /// <param name="row"></param>
        /// <param name="value"></param>
        public void AddString(int column, int row, string value)
        {
            rowCellData.Add(new CellStruct(value));
            rowCoords.Add(new CellCoords(column, row));
        }

        /// <summary>
        /// Add a numerical value to a cell
        /// </summary>
        /// <param name="column"></param>
        /// <param name="row"></param>
        /// <param name="value"></param>
        public void AddNumber(int column, int row, double value)
        {
            rowCellData.Add(new CellStruct(value));
            rowCoords.Add(new CellCoords(column, row));
        }
    }
}
