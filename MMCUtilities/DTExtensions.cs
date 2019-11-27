using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Text;

namespace MMCSirUtilities
{
    public static class DTExtensions
    {
        /// <summary>
        /// Function used to determined wich SQL Type will be used in the SQL Table Type
        /// </summary>
        /// <param name="dtValueType">Column Type (String, Int, Bool, DateTime)</param>
        /// <returns></returns>
        private static string SQLDataType(Type dtValueType)
        {
            string valueType = string.Empty;
            switch (Type.GetTypeCode(dtValueType))
            {
                case TypeCode.String:
                    valueType = "NVARCHAR(MAX)";
                    break;

                case TypeCode.Int16:
                case TypeCode.Int32:
                case TypeCode.Int64:
                    valueType = "INT";
                    break;

                case TypeCode.DateTime:
                    valueType = "DATETIME";
                    break;

                case TypeCode.Boolean:
                    valueType = "BIT";
                    break;

                default:
                    valueType = "NVARCHAR(MAX)";
                    break;
            }
            return valueType.ToUpper();
        }

        /// <summary>
        /// Get an script to create a SQL Table Data Type
        /// </summary>
        /// <param name="typeName">Provide the Name to be asigned to Table Type</param>
        /// <returns></returns>
        public static string ToSQLScriptTableDataType(this DataTable dataTableObject, string typeName)
        {
            StringBuilder sbScript = new StringBuilder();
            sbScript.AppendLine($"CREATE TYPE [dbo].[{typeName}] AS TABLE(");
            for (int i = 0; i < dataTableObject.Columns.Count; i++)
            {
                if (i == 0)
                {
                    sbScript.AppendLine($"{dataTableObject.Columns[i].ColumnName.ToUpper()} {SQLDataType(dataTableObject.Columns[i].DataType)} NULL");
                }
                else
                {
                    sbScript.AppendLine($", {dataTableObject.Columns[i].ColumnName.ToUpper()} {SQLDataType(dataTableObject.Columns[i].DataType)} NULL");
                }
            }
            sbScript.AppendLine(")");
            return sbScript.ToString();
        }

        /// <summary>
        /// Transform an IList or List to datatable using the list propertis as column names
        /// </summary>
        /// <typeparam name="T">List Type</typeparam>
        /// <param name="data">List with the data to be transform.</param>
        /// <param name="excludeColumns">This parameter allows excluding properties from the datatable adding in the data annotations the attribute Exclude Property From DataTable.</param>
        /// <returns></returns>
        public static DataTable ToDataTable<T>(this IList<T> data, Boolean excludeColumns = false)
        {
            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(typeof(T));
            DataTable table = new DataTable();
            foreach (PropertyDescriptor prop in properties)
            {
                if (excludeColumns)
                {
                    if (!prop.Attributes.Contains(new ExcludePropertyFromDataTable()))
                        table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
                }
                else
                {
                    table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
                }
            }

            foreach (T item in data)
            {
                DataRow row = table.NewRow();
                foreach (PropertyDescriptor prop in properties)
                {
                    if (excludeColumns)
                    {
                        if (!prop.Attributes.Contains(new ExcludePropertyFromDataTable()))
                            row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                    }
                    else
                    {
                        row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                    }
                }
                table.Rows.Add(row);
            }
            return table;
        }

        /// <summary>
        /// Transform a DataTable to an HTML string.
        /// </summary>
        /// <param name="dtData">Provide the DataTable to get the information</param>
        /// <returns>string with the HTML data.</returns>
        public static string ToHTMLString(this DataTable dtData)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("<table style='width:100%;text-align:left;border-collapse: collapse;'>");

            #region Create table header

            sb.AppendLine("<tr>");
            foreach (DataColumn colunm in dtData.Columns)
            {
                sb.AppendLine($"<th style='border-bottom: 1px solid #ddd;'>{colunm.ColumnName}</th>");
            }
            sb.AppendLine("</tr>");

            #endregion Create table header

            #region Create table body

            foreach (DataRow row in dtData.Rows)
            {
                sb.AppendLine("<tr>");
                for (int i = 0; i < dtData.Columns.Count; i++)
                {
                    sb.AppendLine($"<td style='border-bottom: 1px solid #ddd;text-align: left; padding: 5px;'>{row[i]}</td>");
                }
                sb.AppendLine("</tr>");
            }

            #endregion Create table body

            sb.AppendLine("</table>");
            return sb.ToString();
        }
    }
}