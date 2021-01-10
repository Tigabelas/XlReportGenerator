using OfficeOpenXml;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace XlReportGenerator
{
    public class FileExcel
    {
        public static List<T> ReadWorksheet<T>(String fileName, String sheetName, String cellRange)
             where T : new()
        {
            List<T> result = new List<T>();

            if (!String.IsNullOrWhiteSpace(fileName) &&
                !String.IsNullOrWhiteSpace(sheetName) &&
                !String.IsNullOrWhiteSpace(cellRange))
            {
                try
                {
                    if (File.Exists(fileName))
                    {
                        FileInfo file = new FileInfo(fileName);
                        using (ExcelPackage package = new ExcelPackage(file))
                        {
                            ExcelWorksheet wSheet = package.Workbook.Worksheets[sheetName];
                            if (wSheet != null)
                            {
                                String[] arrCell = cellRange.Split(':');
                                Int32 startRow = 1;
                                Int32 endRow = 0;
                                Int32 startColumn = 1;
                                Int32 endColumn = 0;

                                //single cell
                                startColumn = GetColumnIndex(Regex.Replace(arrCell[0], @"[0-9\s]", String.Empty));
                                endColumn = startColumn;
                                startRow = Int32.Parse(Regex.Replace(arrCell[0], @"[A-Za-z\s]", String.Empty));
                                endRow = startRow;

                                if (arrCell.Count() == 2)
                                {
                                    endColumn = GetColumnIndex(Regex.Replace(arrCell[1], @"[0-9\s]", String.Empty));
                                    endRow = Int32.Parse(Regex.Replace(arrCell[1], @"[A-Za-z\s]", String.Empty));
                                }

                                Dictionary<String, Int32> mappedColumnToPropertyIndex = null;
                                ArrayList columnContents = new ArrayList();
                                int curRow = startRow;
                                bool isAllColumnNull = false;

                                while (!isAllColumnNull)
                                {
                                    try
                                    {
                                        T objT = new T();

                                        isAllColumnNull = true;
                                        for (var curCol = startColumn; curCol <= endColumn; curCol++)
                                        {
                                            String cellValue = null;

                                            if (wSheet.Cells[curRow, curCol].Value != null)
                                                cellValue = wSheet.Cells[curRow, curCol].Value.ToString();


                                            if (curRow != startRow && mappedColumnToPropertyIndex != null
                                                && mappedColumnToPropertyIndex.ContainsKey(curCol.ToString()) && cellValue != null)
                                            {
                                                objT.GetType().GetProperties()[mappedColumnToPropertyIndex[curCol.ToString()]].SetValue(objT, cellValue);
                                                isAllColumnNull = false;
                                            }

                                            if (curRow == startRow)
                                            {
                                                isAllColumnNull = false;
                                                columnContents.Add(cellValue);
                                            }
                                        }

                                        if (curRow == startRow)
                                        {
                                            mappedColumnToPropertyIndex = MapHeaderToPropertyIndex(objT, columnContents);
                                        }
                                        else
                                        {
                                            if (!isAllColumnNull)
                                                result.Add(objT);
                                        }

                                        curRow++;

                                    }
                                    catch
                                    {
                                        isAllColumnNull = true;
                                    }

                                } // end while
                            }
                        }
                    } // end if file exist
                }
                catch
                {
                    throw;
                }
            }


            return result;
        }

        /// <summary>
        /// Convert column name (in alphabet) into column number
        /// </summary>
        /// <param name="columnName"></param>
        /// <returns></returns>
        private static Int32 GetColumnIndex(String columnName)
        {
            Int32 result = 0;

            Int32 columnNameLength = columnName.Length;

            if (!String.IsNullOrWhiteSpace(columnName))
            {
                for (Int32 i = 0; i < columnNameLength; i++)
                {
                    if (i == 0)
                        result += (Convert.ToByte(columnName[columnNameLength - i - 1]) - 64);
                    else
                        result += 26 * i * (Convert.ToByte(columnName[columnNameLength - i - 1]) - 64);
                }
            }

            return result;
        }
        /// <summary>
        /// MapHeaderToPropertyIndex
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="colNames"></param>
        /// <returns></returns>
        private static Dictionary<String, Int32> MapHeaderToPropertyIndex(Object obj, ArrayList colNames)
        {
            Dictionary<String, Int32> result = new Dictionary<String, Int32>();

            for (var i = 0; i < colNames.Count; i++)
            {
                if (colNames[i] != null)
                {
                    for (var j = 0; j < obj.GetType().GetProperties().Count(); j++)  // loop for all property inside the object
                    {
                        String colName = colNames[i].ToString().Replace(" ", String.Empty).Replace(".", String.Empty).Replace("-", String.Empty).Replace("(", String.Empty).Replace(")", String.Empty).Replace("/", String.Empty);

                        if (colName.ToUpper().Equals(obj.GetType().GetProperties()[j].Name.ToUpper()))
                        {
                            result.Add((i + 1).ToString(), j);
                            break;
                        }
                    }
                }
            }

            return result;
        }
    }
}
