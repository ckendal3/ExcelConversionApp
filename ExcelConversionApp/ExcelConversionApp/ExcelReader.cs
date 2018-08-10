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
        public List<NewData> ReadWorkBook(string path, List<CellMap> cellMaps)
        {
            // list of all data to keep
            List<NewData> dataList = new List<NewData>();

            // row data
            NewData newData;

            FileStream file = File.OpenRead(path);
            IWorkbook workbook = new XSSFWorkbook(path);
            ISheet sheet = workbook.GetSheetAt(0);


            IRow tmpRow;
            // for every row (contact) in the sheet
            for (int i = 0; i < sheet.LastRowNum + 1; i++)
            {
                newData = new NewData();

                // temporary row handler
                tmpRow = sheet.GetRow(i);

                // for every map, get the appropriate data
                foreach (CellMap map in cellMaps)
                {
                    // try to get string value
                    try
                    {
                        // add this STRING data
                        newData.stringData.Add(map.ConversionCellId, tmpRow.GetCell(map.ImportedCellId).StringCellValue);
                    }
                    catch (Exception)
                    {
                        // didn't get string value, try numeric data
                        try
                        {

                            // add this NUMERIC data as a long
                            newData.numericData.Add(map.ConversionCellId, Convert.ToInt64(tmpRow.GetCell(map.ImportedCellId).NumericCellValue));
                        }
                        catch (Exception)
                        {

                            // go to next object
                            continue;
                        }

                        // go to next object
                        continue;
                    }

                }

                dataList.Add(newData);
            }

            return dataList;
        }

    }
}
