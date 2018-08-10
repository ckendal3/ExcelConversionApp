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
        public void CreateWorkBook(string path, string fileName, List<NewData> inData)
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
        public void GenerateRow(ISheet sheet, int rowId, NewData data)
        {
            // For every string data piece - create a corresponding cell
            foreach (KeyValuePair<int, string> stringCell in data.stringData)
            {
                sheet.CreateRow(rowId).CreateCell(stringCell.Key).SetCellType(CellType.String);
                sheet.CreateRow(rowId).CreateCell(stringCell.Key).SetCellValue(stringCell.Value);
            }

            // For every numeric data piece - create a corresponding cell
            foreach (KeyValuePair<int, long> numericCell in data.numericData)
            {
                sheet.CreateRow(rowId).CreateCell(numericCell.Key).SetCellType(CellType.Numeric);
                sheet.CreateRow(rowId).CreateCell(numericCell.Key).SetCellValue(numericCell.Value);
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
