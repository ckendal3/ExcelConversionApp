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
            for (int i = sheet.FirstRowNum; i < sheet.LastRowNum + 1; i++)
            {
                rowData = new RowData();

                // temporary row handler
                tmpRow = sheet.GetRow(i);

                if (tmpRow == null)
                {
                    Console.WriteLine("Row is null");
                    continue;
                }

                // for every mapping, add the data for this row
                for (int j = 0; j < mapArray.Length; j++)
                {
                    if (tmpRow.GetCell(mapArray[j].ImportedCellId) == null)
                    {
                        Console.WriteLine("Cell is null");
                        continue;
                    }

                    if (tmpRow.GetCell(mapArray[j].ImportedCellId).CellType == CellType.String)
                    {
                        rowData.AddString(mapArray[j].ConversionCellId, tmpRow.GetCell(mapArray[j].ImportedCellId).StringCellValue);
                        Console.WriteLine("Value: String");
                    }
                    else if (tmpRow.GetCell(mapArray[j].ImportedCellId).CellType == CellType.Numeric )
                    {
                        rowData.AddNumber(mapArray[j].ConversionCellId, tmpRow.GetCell(mapArray[j].ImportedCellId).NumericCellValue);
                        Console.WriteLine("Value: Numeric");
                    }
                    else
                    {
                        rowData.AddString(mapArray[j].ConversionCellId, "Exception");
                        Console.WriteLine("Value: Exception");
                    }
                }

                dataList.Add(rowData);
            }

            return dataList;
        }

    }
}
