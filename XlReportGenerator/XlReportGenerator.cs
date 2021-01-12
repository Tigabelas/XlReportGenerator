using OfficeOpenXml;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
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

        private static Boolean WriteToCell(ref ExcelWorksheet wBook, Int32 row, Int32 col, Object data, String fieldFormat = "", Boolean isHyperlink = false)
        {
            if (data == null)
                return false;

            Type dataType = data.GetType();

            if (dataType.Equals(typeof(Double))
                                            || dataType.Equals(typeof(Int16))
                                            || dataType.Equals(typeof(Int32))
                                            || dataType.Equals(typeof(Int64))
                                            || dataType.Equals(typeof(Decimal)))
            {
                wBook.Cells[row, col].Value = Convert.ToDouble(data);
            }
            else if (dataType.Equals(typeof(Boolean)))
            {
                wBook.Cells[row, col].Value = (Boolean)data;
            }
            else if (dataType.Equals(typeof(DateTime)))
            {
                if (!String.IsNullOrWhiteSpace(fieldFormat))
                {
                    wBook.Cells[row, col].Value = ((DateTime)data).ToString(fieldFormat);
                }
                else
                {
                    wBook.Cells[row, col].Value = ((DateTime)data).ToString("dd MMM yyyy hh:mm:ss");
                }
            }
            else if (dataType.Equals(typeof(DateTimeOffset)))
            {
                if (!String.IsNullOrWhiteSpace(fieldFormat))
                {
                    wBook.Cells[row, col].Value = ((DateTimeOffset)data).ToString(fieldFormat);
                }
                else
                {
                    wBook.Cells[row, col].Value = ((DateTimeOffset)data).ToString("dd MMM yyyy hh:mm:ss");
                }
            }
            else if (dataType.Equals(typeof(String)))
            {
                if (isHyperlink == false)
                    wBook.Cells[row, col].Value = (String)data;
                else
                {
                    wBook.Cells[row, col].Value = (String)data;
                    try
                    {
                        Uri targetUrl;
                        if (Uri.TryCreate((String)data, UriKind.RelativeOrAbsolute, out targetUrl))
                        {
                            wBook.Cells[row, col].Hyperlink = targetUrl;
                        }
                    }
                    catch
                    {

                    }

                }
            }

            return true;
        }

        private static Int32 WriteDataToSheet(Object data, ref ExcelWorksheet wBook, String sheetName, Int32 startColumn, Int32 startRow, out Int32 maxRow)
        {
            Int32 result = startColumn; // row affected 
            Int32 currentRow = startRow;
            Int32 currentColumn = startColumn;

            maxRow = 0;

            // check whether the sheet is already exis or not
            if (data != null && wBook != null)
            {
                if (data.GetType().GetMethod("GetEnumerator") != null)
                {
                    IEnumerator<Object> enumerator = (IEnumerator<Object>)data.GetType().GetMethod("GetEnumerator").Invoke(data, null);

                    Int32 count = (Int32)data.GetType().GetProperty("Count").GetValue(data);

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
                    PropertyInfo[] propInfos = data.GetType().GetProperties();

                    if (propInfos != null && propInfos.Count() > 0)
                    {
                        foreach (var propInfo in propInfos)
                        {
                            ColumnName colNameAttribute = propInfo.GetCustomAttribute<ColumnName>(true);
                            Skipped skippedAttribute = propInfo.GetCustomAttribute<Skipped>(true);
                            String fieldFormat = (propInfo.GetCustomAttribute<DateFormat>(true) != null) ? propInfo.GetCustomAttribute<DateFormat>(true).Format : "";
                            Boolean isHyperlink = (propInfo.GetCustomAttribute<HyperlinkFormat>(true) != null) ? propInfo.GetCustomAttribute<HyperlinkFormat>(true).IsHyperlink : false;
                            String columnName = "";
                            String fieldName = propInfo.Name;
                            Type fieldType = propInfo.PropertyType;
                            Object fieldValue = propInfo.GetValue(data);
                            Type dataType = data.GetType();

                            if (skippedAttribute == null || (skippedAttribute != null && skippedAttribute.IsSkipped(sheetName)))
                            {
                                bool isTypeNullable = false;

                                if (Nullable.GetUnderlyingType(fieldType) != null)
                                {
                                    // It's nullable
                                    isTypeNullable = true;
                                }

                                if (!SystemTypes.Contains(fieldType) && fieldType.IsClass && !isTypeNullable)
                                {
                                    Int32 curMaxRow = 0;
                                    currentColumn = WriteDataToSheet(fieldValue, ref wBook, sheetName, currentColumn, currentRow, out curMaxRow);

                                    if (curMaxRow > maxRow)
                                        maxRow = curMaxRow;
                                }
                                else
                                {
                                    // Get Attributes[
                                    if (colNameAttribute != null)
                                        columnName = colNameAttribute.Name;
                                    else
                                        columnName = propInfo.Name;

                                    // write header if this is the first row
                                    if (currentRow == 1)
                                    {
                                        WriteToCell(ref wBook, currentRow, currentColumn, columnName, fieldFormat, isHyperlink);
                                    }

                                    // write the data
                                    if (currentRow == 1)
                                        WriteToCell(ref wBook, currentRow + 1, currentColumn, fieldValue, fieldFormat, isHyperlink);
                                    else
                                        WriteToCell(ref wBook, currentRow, currentColumn, fieldValue, fieldFormat, isHyperlink);

                                    currentColumn++;
                                }
                            }
                        }

                        if (currentRow == 1)
                            currentRow = 2;  // we must write header and the first data row

                        if (currentRow > maxRow)
                            maxRow = currentRow;
                    }
                }


                result = currentColumn;
            }

            return result;
        }
        #endregion

        #region "Public Methods"
        /// <summary>
        /// To write from data wrapper to excel file, you can either use template file or generate to new file
        /// </summary>
        /// <param name="data">IEnumerable object to be written</param>
        /// <param name="tempGeneratedFolder">The folder which generated file will be placed.</param>
        /// <param name="sheetName">Sheet name where data should be written</param>
        /// <param name="generatedFileName">Generated file name, will be out to this variable</param>
        /// <param name="workbookTitle">Title for the work book.</param>
        /// <param name="workbookAuthor">The author of the file</param>
        /// <param name="workbookSubject">The subject work book</param>
        /// <param name="workbookKeywords">The keyword</param>
        /// <param name="fromTemplateFileName">Fill with the template file name, if you want to generate from another file</param>
        /// <param name="templatePassword">Fill with the template password if any</param>
        /// <param name="excelType">Output excel file type</param>
        /// <returns></returns>
        public static Int32 Generate(Object data,
            string tempGeneratedFolder,
            string sheetName,
            out string generatedFileName,
            string workbookTitle = "",
            string workbookAuthor = "",
            string workbookSubject = "",
            string workbookKeywords = "",
            string fromTemplateFileName = "",
            string templatePassword = "",
            string outputFileName = "",
            EnumExcelType excelType = EnumExcelType.XLSX)
        {
            Int32 result = 0;
            String fullFileName;
            generatedFileName = String.Empty;

            if (ValidateTempGeneratedFolder(tempGeneratedFolder))
            {
                if (String.IsNullOrWhiteSpace(outputFileName))
                    generatedFileName = GenerateRandomFileName();
                else
                    generatedFileName = outputFileName;

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

                    fullFileName = Path.Combine(tempGeneratedFolder, generatedFileName);

                    FileInfo fileOutput = new FileInfo(fullFileName.ToString());

                    if (fileOutput.Exists)
                    {
                        fileOutput.Delete();
                        fileOutput = new FileInfo(fullFileName.ToString());
                    }

                    Int32 maxRow = 0;

                    if (String.IsNullOrWhiteSpace(fromTemplateFileName))
                    {
                        using (ExcelPackage package = new ExcelPackage())
                        {
                            package.Workbook.Properties.Title = workbookTitle;
                            package.Workbook.Properties.Author = workbookAuthor;
                            package.Workbook.Properties.Subject = workbookSubject;
                            package.Workbook.Properties.Keywords = workbookKeywords;

                            // add a new worksheet to the empty workbook
                            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(sheetName);
                            WriteDataToSheet(data, ref worksheet, sheetName, 1, 1, out maxRow);

                            package.SaveAs(fileOutput);

                            result = 1;
                        }
                    }
                    else
                    {
                        FileInfo fileTemplate = new FileInfo(fromTemplateFileName);
                        if (!fileTemplate.Exists)
                        {
                            throw new Exception("File template doesn't exists.");
                        }

                        using (ExcelPackage package = new ExcelPackage(fileTemplate))
                        {
                            package.Workbook.Properties.Title = workbookTitle;
                            package.Workbook.Properties.Author = workbookAuthor;
                            package.Workbook.Properties.Subject = workbookSubject;
                            package.Workbook.Properties.Keywords = workbookKeywords;

                            // add a new worksheet to the empty workbook
                            ExcelWorksheet worksheet = null;
                            if (package.Workbook.Worksheets.Any(x => x.Name.Equals(sheetName)))
                                worksheet = package.Workbook.Worksheets[sheetName];
                            else
                                worksheet = package.Workbook.Worksheets.Add(sheetName);

                            WriteDataToSheet(data, ref worksheet, sheetName, 1, 1, out maxRow);
                            package.SaveAs(fileOutput);

                            result = 1;
                        }
                    }

                }
                catch
                {
                    throw;
                }
            }

            return result;
        }



        public static Int32 GenerateEx(IEnumerable<ObjectDataWrapper> datas,
            string tempGeneratedFolder,
            out string generatedFileName,
            string workbookTitle,
            string workbookAuthor,
            string workbookSubject,
            string workbookKeywords,
            string fromTemplateFileName = "",
            string templatePassword = "",
            string outputFileName = "",
            EnumExcelType excelType = EnumExcelType.XLSX)
        {
            Int32 result = 0;
            String fullFileName;
            generatedFileName = String.Empty;

            if (datas == null)
                return 0;

            if (ValidateTempGeneratedFolder(tempGeneratedFolder))
            {
                if (String.IsNullOrWhiteSpace(outputFileName))
                    generatedFileName = GenerateRandomFileName();
                else
                    generatedFileName = outputFileName;

                //Create new excel file based on Excel 97/2003
                try
                {
                    if (excelType == EnumExcelType.XLSX)
                        generatedFileName += ".xlsx";
                    else
                        generatedFileName += ".xls";

                    fullFileName = Path.Combine(tempGeneratedFolder, generatedFileName);

                    FileInfo fileOutput = new FileInfo(fullFileName.ToString());

                    if (fileOutput.Exists)
                    {
                        fileOutput.Delete();
                        fileOutput = new FileInfo(fullFileName.ToString());
                    }

                    Int32 maxRow = 0;

                    if (String.IsNullOrWhiteSpace(fromTemplateFileName))
                    {
                        using (ExcelPackage package = new ExcelPackage())
                        {
                            package.Workbook.Properties.Title = workbookTitle;
                            package.Workbook.Properties.Author = workbookAuthor;
                            package.Workbook.Properties.Subject = workbookSubject;
                            package.Workbook.Properties.Keywords = workbookKeywords;

                            foreach (var data in datas)
                            {
                                // add a new worksheet to the empty workbook
                                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(data.SheetName);
                                WriteDataToSheet(data.Data, ref worksheet, data.SheetName, 1, 1, out maxRow);
                            }

                            package.SaveAs(fileOutput);

                            result = 1;
                        }
                    }
                    else
                    {
                        FileInfo fileTemplate = new FileInfo(fromTemplateFileName);
                        if (!fileTemplate.Exists)
                        {
                            throw new Exception("File template doesn't exists.");
                        }

                        using (ExcelPackage package = new ExcelPackage(fileTemplate))
                        {
                            package.Workbook.Properties.Title = workbookTitle;
                            package.Workbook.Properties.Author = workbookAuthor;
                            package.Workbook.Properties.Subject = workbookSubject;
                            package.Workbook.Properties.Keywords = workbookKeywords;

                            foreach (var data in datas)
                            {
                                // add a new worksheet to the empty workbook
                                ExcelWorksheet worksheet = null;
                                if (package.Workbook.Worksheets.Any(x => x.Name.Equals(data.SheetName)))
                                    worksheet = package.Workbook.Worksheets[data.SheetName];
                                else
                                    worksheet = package.Workbook.Worksheets.Add(data.SheetName);

                                WriteDataToSheet(data.Data, ref worksheet, data.SheetName, 1, 1, out maxRow);
                            }

                            package.SaveAs(fileOutput);

                            result = 1;
                        }
                    }

                }
                catch
                {
                    throw;
                }
            }

            return result;
        }


        public static Int32 WriteRawToFile(IEnumerable<RawDataWrapper> datas,
            string sheetName,
            string tempGeneratedFolder,
            out string generatedFileName,
            string workbookTitle,
            string workbookAuthor,
            string workbookSubject,
            string workbookKeywords,
            string fromTemplateFileName = "",
            string templatePassword = "",
            string outputFileName = "",
            int startCol = 1,
            int startRow = 1,
            EnumExcelType excelType = EnumExcelType.XLSX)
        {
            Int32 result = 0;
            String fullFileName;
            generatedFileName = String.Empty;

            if (datas == null)
                return 0;

            if (ValidateTempGeneratedFolder(tempGeneratedFolder))
            {
                if (String.IsNullOrWhiteSpace(outputFileName))
                    generatedFileName = GenerateRandomFileName();
                else
                    generatedFileName = outputFileName;

                //Create new excel file based on Excel 97/2003
                try
                {
                    if (excelType == EnumExcelType.XLSX)
                        generatedFileName += ".xlsx";
                    else
                        generatedFileName += ".xls";

                    fullFileName = Path.Combine(tempGeneratedFolder, generatedFileName);

                    FileInfo fileOutput = new FileInfo(fullFileName.ToString());

                    if (fileOutput.Exists)
                    {
                        fileOutput.Delete();
                        fileOutput = new FileInfo(fullFileName.ToString());
                    }

                    Int32 maxRow = 0;

                    if (String.IsNullOrWhiteSpace(fromTemplateFileName))
                    {
                        using (ExcelPackage package = new ExcelPackage())
                        {
                            package.Workbook.Properties.Title = workbookTitle;
                            package.Workbook.Properties.Author = workbookAuthor;
                            package.Workbook.Properties.Subject = workbookSubject;
                            package.Workbook.Properties.Keywords = workbookKeywords;

                            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(sheetName);

                            foreach (var data in datas)
                            {
                                // add a new worksheet to the empty workbook
                                Int32 col = GetColumnIndex(Regex.Replace(data.Cell, @"[0-9\s]", String.Empty));
                                Int32 row = Int32.Parse(Regex.Replace(data.Cell, @"[A-Za-z\s]", String.Empty));
                                WriteToCell(ref worksheet, row, col, data.Value, "");
                            }

                            package.SaveAs(fileOutput);

                            result = 1;
                        }
                    }
                    else
                    {
                        FileInfo fileTemplate = new FileInfo(fromTemplateFileName);
                        if (!fileTemplate.Exists)
                        {
                            throw new Exception("File template doesn't exists.");
                        }

                        using (ExcelPackage package = new ExcelPackage(fileTemplate))
                        {
                            package.Workbook.Properties.Title = workbookTitle;
                            package.Workbook.Properties.Author = workbookAuthor;
                            package.Workbook.Properties.Subject = workbookSubject;
                            package.Workbook.Properties.Keywords = workbookKeywords;

                            // add a new worksheet to the empty workbook
                            ExcelWorksheet worksheet = null;
                            if (package.Workbook.Worksheets.Any(x => x.Name.Equals(sheetName)))
                                worksheet = package.Workbook.Worksheets[sheetName];
                            else
                                worksheet = package.Workbook.Worksheets.Add(sheetName);

                            foreach (var data in datas)
                            {

                                Int32 col = GetColumnIndex(Regex.Replace(data.Cell, @"[0-9\s]", String.Empty));
                                Int32 row = Int32.Parse(Regex.Replace(data.Cell, @"[A-Za-z\s]", String.Empty));
                                WriteToCell(ref worksheet, row, col, data.Value, "");
                            }

                            package.SaveAs(fileOutput);

                            result = 1;
                        }
                    }

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


        #endregion

    }
}
