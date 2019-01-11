using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelConversionApp.Cell
{
    public struct CellStruct
    {
        /// <summary>
        /// This is the cell type used for determing what type of information is stored
        /// </summary>
        ECellType cellType;

        /// <summary>
        /// The cell's string value
        /// </summary>
        string stringValue;

        /// <summary>
        /// The cell's numerical value
        /// </summary>
        double numericalValue;

        /// <summary>
        /// Initialize the Cell's String Value
        /// </summary>
        /// <param name="value"></param>
        public CellStruct(string value)
        {
            cellType = ECellType.String;
            stringValue = value;

            numericalValue = -1;
        }

        /// <summary>
        /// Initialize the Cell's Numerical Value
        /// </summary>
        /// <param name="value"></param>
        public CellStruct(double value)
        {
            cellType = ECellType.Numerical;
            numericalValue = value;

            stringValue = "";
        }

    }

    public enum ECellType
    {
        Numerical,
        String,
        Special
    };
}
