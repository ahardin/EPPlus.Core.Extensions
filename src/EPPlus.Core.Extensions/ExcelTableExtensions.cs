﻿
using OfficeOpenXml;
using OfficeOpenXml.Table;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace EPPlus.Core.Extensions
{
    /// <summary>
    /// Class holds extensions on ExcelTable object
    /// </summary>
    public static class ExcelTableExtensions
    {
        /// <summary>
        /// Returns given Excel table data bounds with regards to header and totals row visibility
        /// </summary>
        /// <param name="excelTable">Extended object</param>
        /// <returns>Address range</returns>
        public static ExcelAddress GetDataBounds(this ExcelTable excelTable)
        {
            return new ExcelAddress(
                excelTable.Address.Start.Row + (excelTable.ShowHeader ? 1 : 0),
                excelTable.Address.Start.Column,
                excelTable.Address.End.Row - (excelTable.ShowTotal ? 1 : 0),
                excelTable.Address.End.Column
            );
        }

        /// <summary>
        /// Validates the Excel table against the generating type.
        /// </summary>
        /// <typeparam name="T">Generating class type</typeparam>
        /// <param name="excelTable">Extended object</param>
        /// <returns>An enumerable of <see cref="ExcelTableConvertExceptionArgs"/> containing </returns>
        public static IEnumerable<ExcelTableConvertExceptionArgs> Validate<T>(this ExcelTable excelTable) where T : class, new()
        {
            IList mapping = PrepareMappings<T>(excelTable);
            var result = new LinkedList<ExcelTableConvertExceptionArgs>();

            ExcelAddress bounds = excelTable.GetDataBounds();

            var item = (T)Activator.CreateInstance(typeof(T));

            // Parse table
            for (int row = bounds.Start.Row; row <= bounds.End.Row; row++)
            {
                foreach (KeyValuePair<int, PropertyInfo> map in mapping)
                {
                    object cell = excelTable.WorkSheet.Cells[row, map.Key + excelTable.Address.Start.Column].Value;

                    PropertyInfo property = map.Value;

                    try
                    {
                        TrySetProperty(item, property, cell);
                    }
                    catch
                    {
                        result.AddLast(
                            new ExcelTableConvertExceptionArgs
                            {
                                ColumnName = excelTable.Columns[map.Key].Name,
                                ExpectedType = property.PropertyType,
                                PropertyName = property.Name,
                                CellValue = cell,
                                CellAddress = new ExcelCellAddress(row, map.Key + excelTable.Address.Start.Column)
                            });
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Generic extension method yielding objects of specified type from Excel table.
        /// </summary>
        /// <remarks>Exceptions are not catched. It works on all or nothing basis. 
        /// Only primitives and enums are supported as property.
        /// Currently supports only tables with header.</remarks>
        /// <typeparam name="T">Type to map to. Type should be a class and should have parameterless constructor.</typeparam>
        /// <param name="excelTable">ExcelTable object to fetch</param>
        /// <param name="skipCastErrors">Determines how the method should handle exceptions when casting cell value to property type. 
        /// If this is true, invalid casts are silently skipped, otherwise any error will cause method to fail with exception.</param>
        /// <returns>An enumerable of the generating type</returns>
        public static IEnumerable<T> AsEnumerable<T>(this ExcelTable excelTable, bool skipCastErrors = false) where T : class, new()
        {
            IList mapping = PrepareMappings<T>(excelTable);

            ExcelAddress bounds = excelTable.GetDataBounds();

            // Parse table
            for (int row = bounds.Start.Row; row <= bounds.End.Row; row++)
            {
                var item = (T)Activator.CreateInstance(typeof(T));

                foreach (KeyValuePair<int, PropertyInfo> map in mapping)
                {
                    object cell = excelTable.WorkSheet.Cells[row, map.Key + excelTable.Address.Start.Column].Value;

                    PropertyInfo property = map.Value;

                    try
                    {
                        TrySetProperty(item, property, cell);
                    }
                    catch (Exception ex)
                    {
                        if (!skipCastErrors)
                            throw new ExcelTableConvertException(
                                "Cell casting error occures",
                                ex,
                                new ExcelTableConvertExceptionArgs
                                {
                                    ColumnName = excelTable.Columns[map.Key].Name,
                                    ExpectedType = property.PropertyType,
                                    PropertyName = property.Name,
                                    CellValue = cell,
                                    CellAddress = new ExcelCellAddress(row, map.Key + excelTable.Address.Start.Column)
                                }
                            );
                    }
                }

                yield return item;
            }
        }

        /// <summary>
        /// Returns objects of specified type from Excel table as list.
        /// </summary>
        /// <remarks>Exceptions are not catched. It works on all or nothing basis. 
        /// Only primitives and enums are supported as property.
        /// Currently supports only tables with header.</remarks>
        /// <typeparam name="T">Type to map to. Type should be a class and should have parameterless constructor.</typeparam>
        /// <param name="excelTable">ExcelTable object to fetch</param>
        /// <param name="skipCastErrors">Determines how the method should handle exceptions when casting cell value to property type. 
        /// If this is true, invalid casts are silently skipped, otherwise any error will cause method to fail with exception.</param>
        /// <returns>An enumerable of the generating type</returns>
        public static IList<T> ToList<T>(this ExcelTable excelTable, bool skipCastErrors = false) where T : class, new()
        {
            return AsEnumerable<T>(excelTable, skipCastErrors).ToList();
        }

        /// <summary>
        /// Prepares mapping using the type and the attributes decorating its properties
        /// </summary>
        /// <typeparam name="T">Type to parse</typeparam>
        /// <param name="excelTable">ExcelTable to get columns from</param>
        /// <returns>A list of mappings from column index to property</returns>
        private static IList PrepareMappings<T>(ExcelTable excelTable)
        {
            IList mapping = new List<KeyValuePair<int, PropertyInfo>>();

            PropertyInfo[] propInfo = typeof(T).GetProperties(BindingFlags.Instance | BindingFlags.Public);

            // Build property-table column mapping
            foreach (PropertyInfo property in propInfo)
            {
                var mappingAttribute = (ExcelTableColumnAttribute)property.GetCustomAttributes(typeof(ExcelTableColumnAttribute), true).FirstOrDefault();
                if (mappingAttribute != null)
                {
                    int col = -1;
                    // There is no case when both column name and index is specified since this is excluded by the attribute
                    // Neither index, nor column name is specified, use property name
                    if (mappingAttribute.ColumnIndex == 0 && string.IsNullOrWhiteSpace(mappingAttribute.ColumnName))
                    {
                        col = excelTable.Columns[property.Name].Position;
                    }

                    // Column index was specified
                    if (mappingAttribute.ColumnIndex > 0)
                    {
                        col = excelTable.Columns[mappingAttribute.ColumnIndex - 1].Position;
                    }

                    // Column name was specified
                    if (!string.IsNullOrWhiteSpace(mappingAttribute.ColumnName))
                    {
                        col = excelTable.Columns[mappingAttribute.ColumnName].Position;
                    }

                    if (col == -1)
                    {
                        throw new ArgumentException("Should never get here, but I can not identify column.");
                    }

                    mapping.Add(new KeyValuePair<int, PropertyInfo>(col, property));
                }
            }

            return mapping;
        }

        /// <summary>
        /// Tries to set property of item
        /// </summary>
        /// <param name="item">target object</param>
        /// <param name="property">property to be set</param>
        /// <param name="cell">cell value</param>
        private static void TrySetProperty(object item, PropertyInfo property, object cell)
        {
            Type type = property.PropertyType;
            Type itemType = item.GetType();

            // If type is nullable, get base type instead
            if (property.PropertyType.IsNullable())
            {
                if (cell == null) return; // If it is nullable, and we have null we should not waste time

                type = type.GetGenericArguments()[0];
            }

            if (type == typeof(string))
            {
                itemType.InvokeMember(
                    property.Name,
                    BindingFlags.Public | BindingFlags.Instance | BindingFlags.SetProperty,
                    null,
                    item,
                    new object[] { cell?.ToString() });

                return;
            }

            if (type == typeof(DateTime))
            {
                DateTime d = DateTime.Parse(cell.ToString());

                itemType.InvokeMember(
                    property.Name,
                    BindingFlags.Public | BindingFlags.Instance | BindingFlags.SetProperty,
                    null,
                    item,
                    new object[] { d });

                return;
            }

            if (type == typeof(bool))
            {
                itemType.InvokeMember(
                    property.Name,
                    BindingFlags.Public | BindingFlags.Instance | BindingFlags.SetProperty,
                    null,
                    item,
                    new object[] { cell });

                return;
            }

            if (type.IsEnum)
            {
                if (cell.GetType() == typeof(string)) // Support Enum conversion from string...
                {
                    itemType.InvokeMember(
                        property.Name,
                        BindingFlags.Public | BindingFlags.Instance | BindingFlags.SetProperty,
                        null,
                        item,
                        new object[] { Enum.Parse(type, cell.ToString(), true) });
                }
                else // ...and numeric cell value
                {
                    Type underType = type.GetEnumUnderlyingType();

                    itemType.InvokeMember(
                        property.Name,
                        BindingFlags.Public | BindingFlags.Instance | BindingFlags.SetProperty,
                        null,
                        item,
                        new object[] { Enum.ToObject(type, Convert.ChangeType(cell, underType)) });
                }

                return;
            }

            if (type.IsNumeric())
            {
                itemType.InvokeMember(
                    property.Name,
                    BindingFlags.Public | BindingFlags.Instance | BindingFlags.SetProperty,
                    null,
                    item,
                    new object[] { Convert.ChangeType(cell, type) });
            }
        }
    }
}
