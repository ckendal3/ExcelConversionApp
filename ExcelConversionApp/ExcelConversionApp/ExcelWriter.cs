using System;
using System.Collections.Generic;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;

namespace ExcelConversionApp
{
    /// <summary>
    /// Creates the new file and populates it with the directed information.
    /// </summary>
    public class ExcelWriter
    {
        public void CreateWorkBook(string path, string fileName, List<RowData> inData)
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet s1 = workbook.CreateSheet("Sheet1");

            // For every element of data - create a new row
            for (int i = 0; i < inData.Count; i++)
            {
                GenerateRow(s1, i, inData[i]);
            }

            using (var fs = File.Create(CreateSavePath(path, fileName) + ".xlsx"))
            {
                workbook.Write(fs);
                fs.Close();
            }
        }

        /// <summary>
        /// Generate a row based on the input data
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rowId"></param>
        /// <param name="data"></param>
        public void GenerateRow(ISheet sheet, int rowId, RowData data)
        {
            ICell tmpCell;
            IRow tmpRow = sheet.CreateRow(rowId);


            // For every string data piece - create a corresponding cell
            foreach (KeyValuePair<int, string> stringCell in data.stringDict)
            {
                tmpCell = tmpRow.CreateCell(stringCell.Key);
                tmpCell.SetCellType(CellType.String);
                tmpCell.SetCellValue(stringCell.Value);
            }

            
            // For every numeric data piece - create a corresponding cell
            foreach (KeyValuePair<int, long> numericCell in data.numericDict)
            {
                tmpCell = tmpRow.CreateCell(numericCell.Key);
                tmpCell.SetCellType(CellType.Numeric);
                tmpCell.SetCellValue(numericCell.Value);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rowId"></param>
        /// <param name="data"></param>
        public void GenerateRowV2(ISheet sheet, int rowId, RowData data)
        {
            ICell tmpCell;
            IRow tmpRow = sheet.CreateRow(rowId);

            //*********NEEDS TO USE THE LIST OF CellData and CellCoordinates

            Dictionary<int, string> stringCells = data.stringDict;
            Dictionary<int, long> numericCells = data.numericDict;

            for (int i = 0; i < data.numericDict.Count; i++)
            {
                tmpCell = tmpRow.CreateCell(i);
                tmpCell.SetCellType(CellType.String);
                tmpCell.SetCellValue(stringCells[i]);
            }

            for (int i = 0; i < data.numericDict.Count; i++)
            {
                tmpCell = tmpRow.CreateCell(i);
                tmpCell.SetCellType(CellType.Numeric);
                tmpCell.SetCellValue(numericCells[i]);
            }

        }

        /// <summary>
        /// Filter out any unneeded characters in the path
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public string CreateSavePath(string path, string fileName)
        {
            int index = path.LastIndexOf('\\');

            return path.Substring(0, index + 1) + fileName;
        }
    }
}
