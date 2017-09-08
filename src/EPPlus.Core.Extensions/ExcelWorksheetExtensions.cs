using OfficeOpenXml;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace EPPlus.Core.Extensions
{
    public static class ExcelWorksheetExtensions
    {
        /// <summary>
        ///  Returns the data bounds of given worksheet 
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="hasHeaderRow"></param>
        /// <returns>Range address</returns>
        public static ExcelAddress GetDataBounds(this ExcelWorksheet worksheet, bool hasHeaderRow = true)
        {
            return new ExcelAddress(
                worksheet.Dimension.Start.Row + (hasHeaderRow ? 1 : 0),
                worksheet.Dimension.Start.Column,
                worksheet.Dimension.End.Row,
                worksheet.Dimension.End.Column
            );
        }

        /// <summary>
        /// Returns the data cell ranges of given worksheet
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="hasHeaderRow"></param>
        /// <returns>A range of cells</returns>
        public static ExcelRange GetExcelRange(this ExcelWorksheet worksheet, bool hasHeaderRow = true)
        {
            return worksheet.Cells[worksheet.GetDataBounds(hasHeaderRow).Address];
        }

        /// <summary>
        /// Extracts an ExcelTable from given ExcelWorkSheet
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="hasHeaderRow"></param>
        /// <returns>An ExcelTable object</returns>
        public static ExcelTable AsExcelTable(this ExcelWorksheet worksheet, bool hasHeaderRow = true)
        {
            if (worksheet.Tables.Any())
            {
                // Has any table on same addresses
                ExcelAddress dataBounds = worksheet.GetDataBounds(false);
                ExcelTable excelTable = worksheet.Tables.FirstOrDefault(x => x.Address.Address.Equals(dataBounds.Address, StringComparison.InvariantCultureIgnoreCase));
                if (excelTable != null)
                {
                    return excelTable;
                }
            }

            // Table names should be unique
            string tableName = $"{worksheet.Name}-{new Random(Guid.NewGuid().GetHashCode()).Next(9999)}";
            worksheet.Tables.Add(worksheet.GetExcelRange(hasHeaderRow), tableName);
            worksheet.Tables[tableName].ShowHeader = false;
            return worksheet.Tables[tableName];
        }

        /// <summary>
        /// Indicates whether ExcelWorksheet contains any formula or not
        /// </summary>
        /// <param name="worksheet"></param>
        /// <returns>true or false</returns>
        public static bool HasAnyFormula(this ExcelWorksheet worksheet)
        {
            return worksheet.Cells.Any(x => !string.IsNullOrEmpty(x.Formula));
        }

        /// <summary>
        /// Extracts a DataTable from the ExcelWorksheet.
        /// </summary>
        /// <param name="worksheet">The ExcelWorksheet.</param>
        /// <param name="hasHeaderRow">Indicates whether worksheet has a header row or not.</param>
        /// <returns>A DataTable object</returns>
        public static DataTable ToDataTable(this ExcelWorksheet worksheet, bool hasHeaderRow = true)
        {
            ExcelAddress dataBounds = worksheet.GetDataBounds(hasHeaderRow);

            IEnumerable<DataColumn> columns = worksheet.AsExcelTable(!hasHeaderRow).Columns.Select(x => new DataColumn(!hasHeaderRow ? "Column" + x.Id : x.Name));

            var dataTable = new DataTable(worksheet.Name);
            dataTable.Columns.AddRange(columns.ToArray());

            for (int rowIndex = dataBounds.Start.Row; rowIndex <= dataBounds.End.Row; ++rowIndex)
            {
                ExcelRangeBase[] inputRow = worksheet.Cells[rowIndex, dataBounds.Start.Column, rowIndex, dataBounds.End.Column].ToArray();
                DataRow row = dataTable.Rows.Add();

                for (var j = 0; j < inputRow.Length; ++j)
                {
                    row[j] = inputRow[j].Value;
                }
            }

            return dataTable;
        }

        /// <summary>
        /// Generic extension method yielding objects of specified type from excel worksheet.
        /// </summary>
        /// <typeparam name="T">Type to map to. Type should be a class and should have parameterless constructor.</typeparam>
        /// <param name="worksheet">The ExcelWorksheet.</param>
        /// <param name="hasHeaderRow">Indicates whether worksheet has a header row or not.</param>
        /// <param name="skipCastErrors">Determines how the method should handle exceptions when casting cell value to property type. 
        /// If this is true, invalid casts are silently skipped, otherwise any error will cause method to fail with exception.</param>
        /// <returns>An enumerator from given worksheet</returns>
        public static IEnumerable<T> AsEnumerable<T>(this ExcelWorksheet worksheet, bool skipCastErrors = false, bool hasHeaderRow = true) where T : class, new()
        {
            return worksheet.AsExcelTable(hasHeaderRow).AsEnumerable<T>(skipCastErrors);
        }

        /// <summary>
        /// Returns objects of specified type from excel worksheet as list.
        /// </summary>
        /// <typeparam name="T">Type to map to. Type should be a class and should have parameterless constructor.</typeparam>
        /// <param name="worksheet">The ExcelWorksheet.</param>
        /// <param name="hasHeaderRow">Indicates whether worksheet has a header row or not.</param>
        /// <param name="skipCastErrors">Determines how the method should handle exceptions when casting cell value to property type. 
        /// If this is true, invalid casts are silently skipped, otherwise any error will cause method to fail with exception.</param>
        /// <returns>A list of objects from given worksheet</returns>
        public static IList<T> ToList<T>(this ExcelWorksheet worksheet, bool skipCastErrors = false, bool hasHeaderRow = true) where T : class, new()
        {
            return worksheet.AsEnumerable<T>(skipCastErrors, hasHeaderRow).ToList();
        }
    }
}
