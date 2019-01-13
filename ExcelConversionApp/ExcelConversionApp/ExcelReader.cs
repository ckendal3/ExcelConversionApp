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
    public class ExcelReader
    {
        public List<RowData> ReadWorkBook(string path, List<CellMap> cellMaps)
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

            // create an array 
            CellMap[] mapArray = cellMaps.ToArray();

            // row data
            RowData rowData;

            IWorkbook workbook = new XSSFWorkbook(path);
            ISheet sheet = workbook.GetSheetAt(0);


            IRow tmpRow;

            // for every row (contact) in the sheet
            for (int i = 0; i < sheet.LastRowNum + 1; i++)
            {
                Console.WriteLine("Reader is at iteration: " + i);

                rowData = new RowData();

                // temporary row handler
                tmpRow = sheet.GetRow(i);

                // for every mapping, add the data for this row
                for(int j = 0; j < mapArray.Length; j++)
                {
                    if (tmpRow.GetCell(mapArray[i].ImportedCellId).GetType() == typeof(string))
                    {
                        rowData.AddString(mapArray[i].ConversionCellId, tmpRow.GetCell(mapArray[i].ImportedCellId).StringCellValue);
                        Console.WriteLine("Value: String");
                    }
                    else if (tmpRow.GetCell(mapArray[i].ImportedCellId).GetType() == typeof(double))
                    {
                        rowData.AddNumber(mapArray[i].ConversionCellId, tmpRow.GetCell(mapArray[i].ImportedCellId).NumericCellValue);
                        Console.WriteLine("Value: Numeric");
                    }
                    else
                    {
                        rowData.AddString(mapArray[i].ConversionCellId, "Exception");
                        Console.WriteLine("Value: Exception");
                    }
                }

                dataList.Add(rowData);
            }

            return dataList;
        }

    }
}
