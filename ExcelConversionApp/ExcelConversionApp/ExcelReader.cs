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
            // list of all data to keep
            List<RowData> dataList = new List<RowData>();

            // row data
            RowData rowData;

            FileStream file = File.OpenRead(path);
   
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

                // for every map, get the appropriate data
                foreach (CellMap map in cellMaps)
                {
                    // try to get string value
                    try
                    {
                        // add this STRING data
                        rowData.AddString(map.ConversionCellId, tmpRow.GetCell(map.ImportedCellId).StringCellValue);
                    }
                    catch (Exception)
                    {
                        Console.WriteLine("String exception caught");
                        // didn't get string value, try numeric data
                        try
                        {
                            // add this NUMERIC data as a long
                            rowData.AddNumber(map.ConversionCellId, tmpRow.GetCell(map.ImportedCellId).NumericCellValue);
                        }
                        catch (Exception)
                        {

                            // go to next object
                            //newData.stringDict.Add(map.ConversionCellId, "Exception");
                            rowData.AddString(map.ConversionCellId, "Exception");
                        }
                    }   
                }
                dataList.Add(rowData);
            }

            return dataList;
        }

    }
}
