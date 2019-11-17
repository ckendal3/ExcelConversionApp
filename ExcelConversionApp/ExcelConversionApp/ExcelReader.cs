using System;
using System.Collections.Generic;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;

namespace ExcelConversionApp
{
    /// <summary>
    /// Read the information from the spreadsheet to import.
    /// </summary>
    public static class ExcelReader
    {
        public static RowData[] ReadWorkBook(string path, CellMap[] maps)
        {
            try
            {
                FileStream file = File.OpenRead(path);
            }
            catch (Exception)
            {
                Console.WriteLine("File is open, you must close it.");
                return null;
            }

            // list of all data to keep
            List<RowData> dataList = new List<RowData>();

            // row data
            RowData rowData;
            IRow tmpRow;

            IWorkbook workbook = new XSSFWorkbook(path);
            ISheet sheet = workbook.GetSheetAt(0);

            // for every row (contact) in the sheet
            for (int i = sheet.FirstRowNum; i < sheet.LastRowNum + 1; i++)
            {
                rowData = new RowData();

                // temporary row handler
                tmpRow = sheet.GetRow(i);

                ICell cell;

                if (tmpRow == null)
                {
                    Console.WriteLine("Row is null");
                    continue;
                }

                // for every mapping, add the data for this row
                for (int j = 0; j < maps.Length; j++)
                {
                    cell = tmpRow.GetCell(maps[j].ImportedCellId);
                    if (cell == null)
                    {
                        Console.WriteLine("Cell is null");
                        continue;
                    }

                    switch (cell.CellType)
                    {
                        case CellType.String:
                            rowData.AddString(maps[j].ConversionCellId, tmpRow.GetCell(maps[j].ImportedCellId).StringCellValue);
                            Console.WriteLine("Value: String");
                            break;
                        case CellType.Numeric:
                            rowData.AddNumber(maps[j].ConversionCellId, tmpRow.GetCell(maps[j].ImportedCellId).NumericCellValue);
                            Console.WriteLine("Value: Numeric");
                            break;
                        default:
                            rowData.AddString(maps[j].ConversionCellId, "Exception");
                            Console.WriteLine("Value: Exception");
                            break;
                    }
                }

                dataList.Add(rowData);
            }

            return dataList.ToArray();
        }

    }
}
