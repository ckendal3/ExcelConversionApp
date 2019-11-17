using NPOI.SS.UserModel;
using System;

// TODO: Implement efficient CellData Type to prevent so many different checks/exceptions
// ---- This struct is holding way too much data. Should look into interface with self-return value ('dynamic casting')
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

    public interface ICellData<T>  where T : struct
    {
        int Index { get; set; }

        Type GetCellType(out T data);
    }

    public struct CellString : ICellData<CellString>
    {
        public int Index { get; set; }
        public string Value;

        public Type GetCellType(out CellString data)
        {
            data = this;

            return typeof(CellString);
        }
    }

    public struct CellNumeric : ICellData<CellNumeric>
    {
        public int Index { get; set; }
        public long Value;

        public Type GetCellType(out CellNumeric data)
        {
            data = this;

            return typeof(CellNumeric);
        }
    }


}
