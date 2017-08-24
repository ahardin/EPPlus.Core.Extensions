﻿using FluentAssertions;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using System.Collections.Generic;
using System.Data;
using System.Linq;

using Xunit;

namespace EPPlus.Core.Extensions.Tests
{
    public class ExcelWorksheetExtensions_Tests : TestBase
    {
        [Fact]
        public void Test_GetDataBounds_With_Header()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet excelWorksheet = excelPackage.Workbook.Worksheets["TEST5"];

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            ExcelAddress dataBounds = excelWorksheet.GetDataBounds();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            dataBounds.Rows.Should().Be(3);
            dataBounds.Columns.Should().Be(3);
        }

        [Fact]
        public void Test_GetDataBounds_Without_Header()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet excelWorksheet = excelPackage.Workbook.Worksheets["TEST5"];

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            ExcelAddress dataBounds = excelWorksheet.GetDataBounds(false);

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            dataBounds.Rows.Should().Be(4);
            dataBounds.Columns.Should().Be(3);
        }

        [Fact]
        public void Test_GetAsExcelTable_With_Header()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet excelWorksheet = excelPackage.Workbook.Worksheets["TEST5"];

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            ExcelTable excelTable = excelWorksheet.AsExcelTable();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------

            IList<StocksNullable> listOfStocks = excelTable.ToList<StocksNullable>();
            listOfStocks.Count.Should().Be(3);
        }

        [Fact]
        public void Test_GetAsExcelTable_Without_Header()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet excelWorksheet = excelPackage.Workbook.Worksheets["TEST5"];

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            ExcelTable excelTable = excelWorksheet.AsExcelTable(false);

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------

            IList<StocksNullable> listOfStocks = excelTable.ToList<StocksNullable>(true);
            listOfStocks.Count.Should().Be(4);
        }

        [Fact]
        public void Test_ToDataTable_With_Headers()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            DataTable dataTable;

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            dataTable = excelPackage.Workbook.Worksheets["TEST5"].ToDataTable();

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            dataTable.Should().NotBeNull($"{nameof(dataTable)} should not be NULL");
            dataTable.Rows.Count.Should().Be(3, "We have 3 records");
        }

        [Fact]
        public void Test_ToDataTable_Without_Headers()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            DataTable dataTable;

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            dataTable = excelPackage.Workbook.Worksheets["TEST5"].ToDataTable(false);

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            dataTable.Should().NotBeNull($"{nameof(dataTable)} should not be NULL");
            dataTable.Rows.Count.Should().Be(4, "We have 4 records");
        }
        

        [Fact]
        public void Test_Worksheet_AsEnumerable()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet1 = excelPackage.Workbook.Worksheets["TEST4"];
            ExcelWorksheet worksheet2 = excelPackage.Workbook.Worksheets["TEST5"];

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            IEnumerable<StocksNullable> list1 = worksheet1.AsEnumerable<StocksNullable>(true);
            IEnumerable<StocksNullable> list2 = worksheet2.AsEnumerable<StocksNullable>(true, false);

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            list1.Count().Should().Be(4, "Should have four");
            list2.Count().Should().Be(4, "Should have four");
        }


        [Fact]
        public void Test_Worksheet_ToList()
        {
            //-----------------------------------------------------------------------------------------------------------
            // Arrange
            //-----------------------------------------------------------------------------------------------------------
            ExcelWorksheet worksheet1 = excelPackage.Workbook.Worksheets["TEST4"];
            ExcelWorksheet worksheet2 = excelPackage.Workbook.Worksheets["TEST5"];

            //-----------------------------------------------------------------------------------------------------------
            // Act
            //-----------------------------------------------------------------------------------------------------------
            IEnumerable<StocksNullable> list1 = worksheet1.ToList<StocksNullable>(true);
            IEnumerable<StocksNullable> list2 = worksheet2.ToList<StocksNullable>(true, false);

            //-----------------------------------------------------------------------------------------------------------
            // Assert
            //-----------------------------------------------------------------------------------------------------------
            list1.Count().Should().Be(4, "Should have four");
            list2.Count().Should().Be(4, "Should have four");
        }
    }
}
