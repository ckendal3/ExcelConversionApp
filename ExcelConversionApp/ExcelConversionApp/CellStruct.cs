using NPOI.SS.UserModel;

namespace ExcelConversionApp
{
    public struct CellData
    {
        /// <summary>
        /// This is the cell type used for determing what type of information is stored
        /// </summary>
        CellType cellType;

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
        public CellData(string value)
        {
            cellType = CellType.String;
            stringValue = value;

            numericalValue = -1;
        }

        /// <summary>
        /// Initialize the Cell's Numerical Value
        /// </summary>
        /// <param name="value"></param>
        public CellData(double value)
        {
            cellType = CellType.Numeric;
            numericalValue = value;

            stringValue = "";
        }

    }


    /// <summary>
    /// Cell coordinate holder. X = Column, Y = Row
    /// </summary>
    public struct CellCoordinates
    {
        /// <summary>
        /// This holds the column location
        /// </summary>
        int x;

        /// <summary>
        /// This holds the row location
        /// </summary> 
        int y;

        public CellCoordinates(int column, int row)
        {
            x = column;
            y = row;
        }
    }
}
