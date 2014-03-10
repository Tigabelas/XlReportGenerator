using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace XlReportGenerator
{
    public class XlReportGenerator
    {
        #region "Private Variables and Properties"
        private static List<Type> _systemTypes;
        public static List<Type> SystemTypes
        {
            get
            {
                if (_systemTypes == null)
                {
                    _systemTypes = Assembly.GetExecutingAssembly().GetType().Module.Assembly.GetExportedTypes().ToList();
                }
                return _systemTypes;
            }
        }

        #endregion
        #region "Private Methods"
        /// <summary>
        /// Generate random filename
        /// </summary>
        /// <returns></returns>
        private static String GenerateRandomFileName()
        {
            String result = String.Empty;

            result = DateTime.Now.Ticks.ToString();

            return result;
        }

        /// <summary>
        /// Validate the TempGeneratedFolder
        /// </summary>
        /// <returns>True if valid, False is not valid</returns>
        private static Boolean ValidateTempGeneratedFolder(String tempGeneratedFolder)
        {
            Boolean result = false;

            // Check the folder whether the folder exist or not
            if (!String.IsNullOrWhiteSpace(tempGeneratedFolder))
            {
                try
                {
                    if (!Directory.Exists(tempGeneratedFolder))
                    {
                        Directory.CreateDirectory(tempGeneratedFolder);
                    }

                    result = true;
                }
                catch
                {
                    result = false;
                }
            }

            return result;
        }

        private static Boolean WriteToCell(ref ISheet sheet, Int32 row, Int32 column, Object data, String fieldFormat="")
        {
            Boolean result = false;
            Type dataType = data.GetType();

            if (sheet != null && data != null && row > -1 && column > -1)
            {
                // Write the column header
                IRow wRow = sheet.GetRow(row);
                ICell wCell = null;

                if (wRow == null)
                    wRow = sheet.CreateRow(row);

                wCell = wRow.GetCell(column);

                if (wCell == null)
                    wCell = wRow.CreateCell(column);

                if (dataType.Equals(typeof(Double))
                    || dataType.Equals(typeof(Int16))
                    || dataType.Equals(typeof(Int32))
                    || dataType.Equals(typeof(Int64))
                    || dataType.Equals(typeof(Decimal)))
                {
                    wCell.SetCellType(CellType.Numeric);
                    wCell.SetCellValue(Convert.ToDouble(data));
                }
                else if (dataType.Equals(typeof(Boolean)))
                {
                    wCell.SetCellType(CellType.Boolean);
                    wCell.SetCellValue((Boolean)data);
                }
                else if (dataType.Equals(typeof(DateTime)))
                {
                    wCell.SetCellType(CellType.String);
                    if (!String.IsNullOrWhiteSpace(fieldFormat))
                        wCell.SetCellValue(((DateTime)data).ToString(fieldFormat));
                    else
                        wCell.SetCellValue(((DateTime)data).ToString("dd MMM yyyy hh:mm:ss"));
                    
                }
                else if (dataType.Equals(typeof(String)))
                {
                    wCell.SetCellType(CellType.String);
                    wCell.SetCellValue((String)data);
                }

                result = true;
            }

            return result;
        }

        private static Int32 WriteDataToSheet(Object data, ref IWorkbook wBook, String sheetName, Int32 startColumn, Int32 startRow, out Int32 maxRow)
        {
            Int32 result = 0; // row affected 
            Int32 currentRow = startRow;
            Int32 currentColumn = startColumn;

            maxRow = 0;

            // check whether the sheet is already exis or not
            if (data != null && wBook != null && !String.IsNullOrWhiteSpace(sheetName))
            {
                if (wBook.GetSheetIndex(sheetName) < 0)
                    wBook.CreateSheet(sheetName);

                ISheet sheet = wBook.GetSheet(sheetName);

                if (sheet != null)
                {
                    if (data.GetType().GetMethod("GetEnumerator") != null )
                    {
                        IEnumerator enumerator = (IEnumerator) data.GetType().GetMethod("GetEnumerator").Invoke(data, null);
                        Int32 count = (Int32) data.GetType().GetProperty("Count").GetValue(data);

                        Int32 i = 0;
                        
                        while (enumerator.MoveNext())
                        {
                            currentColumn = WriteDataToSheet(enumerator.Current, ref wBook, sheetName, currentColumn, currentRow, out currentRow);

                            if (i < count - 1)
                            {
                                currentRow++;
                                currentColumn = startColumn;
                            }

                            if (currentRow > maxRow)
                                maxRow = currentRow;
                            else
                                currentRow = maxRow;

                            i++;
                        }
                    }
                    else
                    {
                        PropertyInfo [] propInfos = data.GetType().GetProperties();

                        if (propInfos != null && propInfos.Count() > 0)
                        {
                            foreach (var propInfo in propInfos)
                            {
                                ColumnName colNameAttribute = propInfo.GetCustomAttribute<ColumnName>(true);
                                Skipped skippedAttribute = propInfo.GetCustomAttribute<Skipped>(true);
                                String fieldFormat = (propInfo.GetCustomAttribute<DateFormat>(true) != null) ? propInfo.GetCustomAttribute<DateFormat>(true).Format : "";
                                String columnName = "";
                                String fieldName = propInfo.Name;
                                Type fieldType = propInfo.PropertyType;
                                Object fieldValue = propInfo.GetValue(data);

                                if (skippedAttribute == null || (skippedAttribute != null && !skippedAttribute.IsSkipped))
                                {
                                    if (!SystemTypes.Contains(fieldType))
                                    {
                                        currentColumn = WriteDataToSheet(fieldValue, ref wBook, sheetName, currentColumn, currentRow, out maxRow);
                                    }
                                    else
                                    {
                                        // Get Attributes[

                                        if (colNameAttribute != null)
                                            columnName = colNameAttribute.Name;
                                        else
                                            columnName = propInfo.Name;

                                        // write header if this is the first row
                                        if (currentRow == 0)
                                            WriteToCell(ref sheet, currentRow, currentColumn, columnName);

                                        // write the data
                                        if (currentRow == 0)
                                            WriteToCell(ref sheet, currentRow + 1, currentColumn, fieldValue, fieldFormat);
                                        else
                                            WriteToCell(ref sheet, currentRow, currentColumn, fieldValue, fieldFormat);

                                        currentColumn++;

                                    }
                                }
                            }

                            // add current row
                            if (currentRow == 0)
                                currentRow = 1;  // we have write header and the first data row

                            if (currentRow > maxRow)
                                maxRow = currentRow;
                        }
                    }
                    
                }
                result = currentColumn;
            }

            return result;
        }

        #endregion


        #region "Public Methods"
        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="T">Data Type which you want to write to excel file</typeparam>
        /// <param name="data">The data source which will be written to excel file</param>
        /// <returns></returns>
        public static Int32 Generate(Object data, String tempGeneratedFolder, String sheetName, EnumExcelType excelType = EnumExcelType.NONE)
        {
            Int32 result = 0;
            IWorkbook wBook = null;
            StringBuilder fullFileName = new StringBuilder();
            FileMode fsMode = FileMode.CreateNew;
            String generatedFileName = String.Empty;

            if (ValidateTempGeneratedFolder(tempGeneratedFolder))
            {
                generatedFileName = GenerateRandomFileName();
                
                //Create new excel file based on Excel 97/2003
                try
                {
                    //Validate SheetName or generate default sheet name from filename
                    if (String.IsNullOrWhiteSpace(sheetName))
                        sheetName = generatedFileName;

                    if (excelType == EnumExcelType.XLSX)
                        generatedFileName += ".xlsx";
                    else
                        generatedFileName += ".xls";

                    fullFileName.Append(tempGeneratedFolder);
                    fullFileName.Append("\\");
                    fullFileName.Append(generatedFileName);

                    if (File.Exists(fullFileName.ToString()))
                        fsMode = FileMode.Open;
                    else
                        fsMode = FileMode.CreateNew;

                    using (FileStream fs = new FileStream(fullFileName.ToString(), fsMode, FileAccess.ReadWrite))
                    {
                        // Check
                        if (excelType == EnumExcelType.XLSX)
                        {
                            if (fsMode == FileMode.Open)
                                wBook = new XSSFWorkbook(fs);
                            else
                                wBook = new XSSFWorkbook();
                        }
                        else
                        {
                            if (fsMode == FileMode.Open)
                                wBook = new HSSFWorkbook(fs);
                            else
                                wBook = new HSSFWorkbook();
                        }

                        //Write data to sheet
                        Int32 maxRow = 0;
                        WriteDataToSheet(data, ref wBook, sheetName, 0, 0, out maxRow);

                        //Write and close the file.
                        wBook.Write(fs);
                        fs.Close();
                    }
                }
                catch
                {
                    throw;
                }
            }

            return result;
        }

        
        #endregion

    }
}
