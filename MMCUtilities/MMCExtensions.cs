using ADODB;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;

namespace MMCSirUtilities
{
    public static class MMCExtentions
    {
        /// <summary>
        /// Transform a recordser to DataTable
        /// </summary>
        /// <param name="rs">RecordSet to be transformed</param>
        /// <returns>Get the datatable with the information obteined from the recordset</returns>
        public static DataTable RecorsetToDataTable(this Recordset rs)
        {
            DataTable dt = new DataTable();
            OleDbDataAdapter odA = new OleDbDataAdapter();
            odA.Fill(dt, rs);
            return dt;
        }

        /// <summary>
        /// Create and excel file from a DataTable.
        /// </summary>
        /// <param name="dt">DataTable with the information.</param>
        /// <param name="path">Set the path to stored the file.</param>
        /// <param name="fileName">Set the file name without extensions</param>
        /// <param name="includeHeaders">Set if the file needs headers</param>
        public static void CreateExcelFile(DataTable dt, string path, string fileName, bool includeHeaders = true)
        {
            if (!Directory.Exists(path))
            {
                throw new DirectoryNotFoundException("The Directory does not exits or you don't have the right access.");
            }
            string fullFileName = $"{path}\\{fileName}.xlsx";

            FileInfo newExcelFile = new FileInfo(fullFileName);
            if (File.Exists(fullFileName))
            {
                File.Delete(fullFileName);
            }
            using (var package = new ExcelPackage(newExcelFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("sheet1");
                int columnCounter = 1;
                int rowcounter = 1;

                if (includeHeaders)
                {
                    foreach (DataColumn column in dt.Columns)
                    {
                        worksheet.Cells[1, columnCounter].Value = column.ColumnName;
                        columnCounter++;
                    }
                    rowcounter++;
                }

                foreach (DataRow rowData in dt.Rows)
                {
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        worksheet.Cells[rowcounter, i + 1].Value = rowData[dt.Columns[i].ColumnName];
                        if (dt.Columns[i].DataType.FullName == "System.DateTime")
                        {
                            worksheet.Cells[rowcounter, i + 1].Style.Numberformat.Format = "yyyy-MM-dd HH:mm:ss";
                        }
                    }
                    rowcounter++;
                }
                package.Save();
            }
        }

        /// <summary>
        /// Create and excel file from a Generic List.
        /// </summary>
        /// <typeparam name="T">Define the Generic List to Use</typeparam>
        /// <param name="lst">Set Generic List with the information</param>
        /// <param name="path">Set the path to stored the file.</param>
        /// <param name="fileName">Set the file name without extensions</param>
        /// <param name="includeHeaders">Set if the file needs headers</param>
        public static void CreateExcelFile<T>(List<T> lst, string path, string fileName, bool includeHeaders = true)
        {
            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(typeof(T));
            if (!Directory.Exists(path))
            {
                throw new DirectoryNotFoundException("The Directory does not exits or you don't have the right access.");
            }
            string fullFileName = $"{path}\\{fileName}.xlsx";

            FileInfo newExcelFile = new FileInfo(fullFileName);
            if (File.Exists(fullFileName))
            {
                File.Delete(fullFileName);
            }
            using (var package = new ExcelPackage(newExcelFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("sheet1");
                int columnCounter = 1;
                int rowcounter = 1;

                if (includeHeaders)
                {
                    foreach (PropertyDescriptor prop in properties)
                    {
                        worksheet.Cells[1, columnCounter].Value = prop.DisplayName;
                        columnCounter++;
                    }
                    rowcounter++;
                }

                foreach (T item in lst)
                {
                    for (int i = 0; i < properties.Count; i++)
                    {
                        PropertyDescriptor prop = properties[i];
                        worksheet.Cells[rowcounter, i + 1].Value = prop.GetValue(item) ?? DBNull.Value;
                        if (prop.PropertyType.FullName == "System.DateTime")
                        {
                            worksheet.Cells[rowcounter, i + 1].Style.Numberformat.Format = "yyyy-MM-dd HH:mm:ss";
                        }
                    }
                    rowcounter++;
                }
                package.Save();
            }
        }

        /// <summary>
        ///
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="elementsList"></param>
        /// <returns></returns>
        public static DataTable ConvertToDataTable<T>(this List<T> elementsList)
        {
            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(typeof(T));
            DataTable table = new DataTable();
            foreach (PropertyDescriptor prop in properties)
                table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
            foreach (T item in elementsList)
            {
                DataRow row = table.NewRow();
                foreach (PropertyDescriptor prop in properties)
                    row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                table.Rows.Add(row);
            }
            return table;
        }

        /// <summary>
        /// Function to create CSV from a DataTable.
        /// </summary>
        /// <param name="dt">DataTable with all the information</param>
        /// <param name="path">Set the path to stored the File.</param>
        /// <param name="fileName">Set the name for the file without extension.</param>
        /// <param name="includeHeaders">Include headers in the file.</param>
        public static void CreateCSVFile(DataTable dt, string path, string fileName, bool includeHeaders = true)
        {
            StringBuilder sb = new StringBuilder();

            if (!Directory.Exists(path))
            {
                throw new DirectoryNotFoundException("The Directory does not exits or you don't have the right access.");
            }
            string fullFileName = $"{path}\\{fileName}.csv";

            FileInfo newExcelFile = new FileInfo(fullFileName);
            if (File.Exists(fullFileName))
            {
                File.Delete(fullFileName);
            }

            if (includeHeaders)
            {
                string strHeader = string.Empty;
                foreach (DataColumn column in dt.Columns)
                {
                    strHeader += $"{column.ColumnName},";
                }
                strHeader = strHeader.Remove(strHeader.Length - 1, 1);
                sb.AppendLine(strHeader);
            }

            foreach (DataRow rowData in dt.Rows)
            {
                string strLine = string.Empty;
                foreach (DataColumn columnName in dt.Columns)
                {
                    strLine += $"{rowData[columnName.ColumnName]},";
                }
                strLine = strLine.Remove(strLine.Length - 1, 1);
                sb.AppendLine(strLine);
            }
            File.WriteAllText(fullFileName, sb.ToString());
        }

        public static void CreateCSVFile<T>(List<T> lst, string path, string fileName, bool includeHeaders = true)
        {
            StringBuilder sb = new StringBuilder();

            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(typeof(T));
            if (!Directory.Exists(path))
            {
                throw new DirectoryNotFoundException("The Directory does not exits or you don't have the right access.");
            }
            string fullFileName = $"{path}\\{fileName}.csv";

            FileInfo newExcelFile = new FileInfo(fullFileName);
            if (File.Exists(fullFileName))
            {
                File.Delete(fullFileName);
            }

            if (includeHeaders)
            {
                string strHeader = string.Empty;
                foreach (PropertyDescriptor prop in properties)
                {
                    strHeader += $"{prop.DisplayName},";
                }
                strHeader = strHeader.Remove(strHeader.Length - 1, 1);
                sb.AppendLine(strHeader);
            }

            foreach (T item in lst)
            {
                string strLine = string.Empty;
                for (int i = 0; i < properties.Count; i++)
                {
                    PropertyDescriptor prop = properties[i];
                    strLine += $"{ prop.GetValue(item) ?? DBNull.Value},";
                }
                strLine = strLine.Remove(strLine.Length - 1, 1);
                sb.AppendLine(strLine);
            }
            File.WriteAllText(fullFileName, sb.ToString());
        }

        /// <summary>
        /// Get a DataTable from an Excel file only the First Sheet
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="hasHeader"></param>
        /// <returns></returns>
        public static DataTable GetDTExcelFirstSheet(string fileName, bool hasHeader)
        {
            var package = new ExcelPackage(new FileInfo(fileName));
            DataTable dtInformation = new DataTable();

            if (package.Workbook.Worksheets.Count > 0)
            {
                int startRow = 1;
                ExcelWorksheet sheet = package.Workbook.Worksheets.FirstOrDefault();
                if (sheet != null)
                {
                    if (hasHeader)
                    {
                        foreach (var firstRowCell in sheet.Cells[1, 1, 1, sheet.Dimension.End.Column])
                        {
                            dtInformation.Columns.Add(firstRowCell.Text);
                        }
                        startRow = 2;
                    }
                    else
                    {
                        int counter = 1;
                        foreach (var firstRowCell in sheet.Cells[1, 1, 1, sheet.Dimension.End.Column])
                        {
                            dtInformation.Columns.Add($"ColumnName{counter}");
                            counter++;
                        }
                    }

                    for (var rowNumber = startRow; rowNumber <= sheet.Dimension.End.Row; rowNumber++)
                    {
                        var excelRow = sheet.Cells[rowNumber, 1, rowNumber, sheet.Dimension.End.Column];
                        var dtNewRow = dtInformation.NewRow();
                        foreach (var excelCell in excelRow)
                        {
                            dtNewRow[excelCell.Start.Column - 1] = excelCell.Text;
                        }
                        dtInformation.Rows.Add(dtNewRow);
                    }
                }
            }
            return dtInformation;
        }

        /// <summary>
        /// Export a Generic list to Excel in a WebPage.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="lstSource">Source Data</param>
        /// <param name="fileName">Excel File Name without Extension</param>
        /// <param name="includeHeaders">Include the headers in the excel file.</param>
        public static void ExportListToExcel<T>(List<T> lstSource, string fileName, bool includeHeaders = true)
        {
            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(typeof(T));
            using (var package = new ExcelPackage())
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("sheet1");
                int columnCounter = 1;
                int rowcounter = 1;
                if (includeHeaders)
                {
                    foreach (PropertyDescriptor prop in properties)
                    {
                        worksheet.Cells[1, columnCounter].Value = prop.DisplayName;
                        columnCounter++;
                    }
                    rowcounter++;
                }
                foreach (T item in lstSource)
                {
                    for (int i = 0; i < properties.Count; i++)
                    {
                        PropertyDescriptor prop = properties[i];
                        worksheet.Cells[rowcounter, i + 1].Value = prop.GetValue(item) ?? DBNull.Value;
                        if (prop.PropertyType.FullName == "System.DateTime")
                        {
                            worksheet.Cells[rowcounter, i + 1].Style.Numberformat.Format = "yyyy-MM-dd HH:mm:ss";
                        }
                    }
                    rowcounter++;
                }
                using (var memoryStream = new MemoryStream())
                {
                    HttpContext.Current.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                    HttpContext.Current.Response.AddHeader("content-disposition", $"attachment; filename={fileName}.xlsx");
                    package.SaveAs(memoryStream);
                    memoryStream.WriteTo(HttpContext.Current.Response.OutputStream);
                    HttpContext.Current.Response.Flush();
                    HttpContext.Current.Response.End();
                }
            }
        }

        public static void ExportDataTableToExcel(DataTable dtSource, string fileName, bool includeHeaders = true)
        {
            using (var package = new ExcelPackage())
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("sheet1");
                int columnCounter = 1;
                int rowcounter = 1;
                if (includeHeaders)
                {
                    foreach (DataColumn column in dtSource.Columns)
                    {
                        worksheet.Cells[1, columnCounter].Value = column.ColumnName;
                        columnCounter++;
                    }
                    rowcounter++;
                }
                foreach (DataRow rowData in dtSource.Rows)
                {
                    for (int i = 0; i < dtSource.Columns.Count; i++)
                    {
                        worksheet.Cells[rowcounter, i + 1].Value = rowData[dtSource.Columns[i].ColumnName];
                        if (dtSource.Columns[i].DataType.FullName == "System.DateTime")
                        {
                            worksheet.Cells[rowcounter, i + 1].Style.Numberformat.Format = "yyyy-MM-dd HH:mm:ss";
                        }
                    }
                    rowcounter++;
                }
                using (var memoryStream = new MemoryStream())
                {
                    HttpContext.Current.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                    HttpContext.Current.Response.AddHeader("content-disposition", $"attachment; filename={fileName}.xlsx");
                    package.SaveAs(memoryStream);
                    memoryStream.WriteTo(HttpContext.Current.Response.OutputStream);
                    HttpContext.Current.Response.Flush();
                    HttpContext.Current.Response.End();
                }
            }
        }
    }
}