# **EPPlus.Core.Extensions** [![Build status](https://ci.appveyor.com/api/projects/status/cdhoa8m20k2k71ke/branch/master?svg=true)](https://ci.appveyor.com/project/eraydin/epplus-core-extensions/branch/master)

An extensions library for both EPPlus and EPPlus.Core packages to generate and manipulate Excel files easily.

### **Installation** [![NuGet version](https://badge.fury.io/nu/EPPlus.Core.Extensions.svg)](https://badge.fury.io/nu/EPPlus.Core.Extensions)

It's as easy as `PM> Install-Package EPPlus.Core.Extensions` from [nuget](http://nuget.org/packages/EPPlus.Core.Extensions)

### **Dependencies**

**.NET Framework 4.6.1**
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;*EPPlus >= 4.1.0*

**.NET Standard 2.0**
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;*EPPlus.Core >= 1.5.2*

---
### **A Brief Summary of Extensions methods**
---
---
## Extensions on ExcelPackage object
---

#### Method: ExcelPackageExtensions.GetTable(OfficeOpenXml.ExcelPackage,System.String)

 Returns concrete ExcelTable by its name 

Parameters:

|Name | Description |
|-----|------|
|excelPackage: |The ExcelPackage object|
|name: |Name of the table|
**Returns**: ExcelTable object if found, null if not

---
#### Method: ExcelPackageExtensions.GetTables(OfficeOpenXml.ExcelPackage)

 Returns all Excel tables in the opened worksheet 

Parameters:

|Name | Description |
|-----|------|
|excelPackage: |The ExcelPackage object|
**Returns**: Enumeration of ExcelTables

---
#### Method: ExcelPackageExtensions.HasTable(OfficeOpenXml.ExcelPackage,System.String)

 Checks the given table name is in the ExcelPackage or not 

Parameters: 

|Name | Description |
|-----|------|
|excelPackage: |The ExcelPackage object|
|name: |Name of the table|
**Returns**: Result of search as bool

---
#### Method ExcelPackageExtensions.ToDataSet(OfficeOpenXml.ExcelPackage,System.Boolean)

 Extracts a DataSet from the ExcelPackage. 

Parameters:

|Name | Description |
|-----|------|
|excelPackage: |The ExcelPackage.|
|hasHeaderRow: |Indicates whether worksheet has a header row or not.|
**Returns**: DataSet object

---
## Extensions on ExcelTable object
---
#### Method ExcelTableExtensions.GetDataBounds(OfficeOpenXml.Table.ExcelTable)

 Returns given Excel table data bounds with regards to header and totals row visibility 

Parameters:

|Name | Description |
|-----|------|
|excelTable: |Extended object|
**Returns**: Address range


---
#### Method ExcelTableExtensions.Validate<T>(OfficeOpenXml.Table.ExcelTable)

 Validates the Excel table against the generating type. 

Parameters:

|Name | Description |
|-----|------|
|T: |Generating class type|
|table: |Extended object|
**Returns**: An enumerable of [[|T:EPPlus.Core.Extensions.ExcelTableConvertExceptionArgs]] containing 



---
#### Method ExcelTableExtensions.AsEnumerable<T>(OfficeOpenXml.Table.ExcelTable,System.Boolean)

 Generic extension method yielding objects of specified type from Excel table. 

Parameters: 
|Name | Description |
|-----|------|
|T: |Type to map to. Type should be a class and should have parameterless constructor.|
|excelTable: |Table object to fetch|
|skipCastErrors: |Determines how the method should handle exceptions when casting cell value to property type. If this is true, invalid casts are silently skipped, otherwise any error will cause method to fail with exception.|
**Returns**: An enumerable of the generating type



---
#### Method ExcelTableExtensions.ToList<T>(OfficeOpenXml.Table.ExcelTable,System.Boolean)

 Returns objects of specified type from Excel table as list. 

Parameters:

|Name | Description |
|-----|------|
|T: |Type to map to. Type should be a class and should have parameterless constructor.|
|excelTable: |Table object to fetch|
|skipCastErrors: |Determines how the method should handle exceptions when casting cell value to property type. If this is true, invlaid casts are silently skipped, otherwise any error will cause method to fail with exception.|
**Returns**: An enumerable of the generating type


---
## Extensions on ExcelWorksheet object
---
#### Method ExcelWorksheetExtensions.GetDataBounds(OfficeOpenXml.ExcelWorksheet,System.Boolean)

 Returns the data bounds of given worksheet 

Parameters:

|Name | Description |
|-----|------|
|worksheet: ||
|hasHeaderRow: ||
**Returns**: Range address



---
#### Method ExcelWorksheetExtensions.GetExcelRange(OfficeOpenXml.ExcelWorksheet,System.Boolean)

 Returns the data cell ranges of given worksheet

Parameters:

|Name | Description |
|-----|------|
|worksheet: ||
|hasHeaderRow: ||
**Returns**: A range of cells



---
#### Method ExcelWorksheetExtensions.AsExcelTable(OfficeOpenXml.ExcelWorksheet,System.Boolean)

 Extracts an ExcelTable from given ExcelWorkSheet 

Parameters:

|Name | Description |
|-----|------|
|worksheet: ||
|hasHeaderRow: ||
**Returns**: An ExcelTable object



---
#### Method ExcelWorksheetExtensions.HasAnyFormula(OfficeOpenXml.ExcelWorksheet)

 Indicates whether ExcelWorksheet contains any formula or not 

Parameters:

|Name | Description |
|-----|------|
|worksheet: ||
**Returns**: true or false



---
#### Method ExcelWorksheetExtensions.ToDataTable(OfficeOpenXml.ExcelWorksheet,System.Boolean)

 Extracts a DataTable from the ExcelWorksheet. 

Parameters:

|Name | Description |
|-----|------|
|worksheet: |The ExcelWorksheet.|
|hasHeaderRow: |Indicates whether worksheet has a header row or not.|
**Returns**: A DataTable object



---
#### Method ExcelWorksheetExtensions.AsEnumerable<T>(OfficeOpenXml.ExcelWorksheet,System.Boolean,System.Boolean)

 Generic extension method yielding objects of specified type from excel worksheet. 

Parameters:

|Name | Description |
|-----|------|
|T:  |Type to map to. Type should be a class and should have parameterless constructor.|
|worksheet: |The ExcelWorksheet.|
|hasHeaderRow: |Indicates whether worksheet has a header row or not.|
|skipCastErrors: |Determines how the method should handle exceptions when casting cell value to property type. If this is true, invalid casts are silently skipped, otherwise any error will cause method to fail with exception.|
**Returns**: An enumerator from given worksheet



---
#### Method ExcelWorksheetExtensions.ToList<T>(OfficeOpenXml.ExcelWorksheet,System.Boolean,System.Boolean)

 Returns objects of specified type from excel worksheet as list. 

|Name | Description |
|-----|------|
|T:  |Type to map to. Type should be a class and should have parameterless constructor.|
|worksheet: |The ExcelWorksheet.|
|hasHeaderRow: |Indicates whether worksheet has a header row or not.|
|skipCastErrors: |Determines how the method should handle exceptions when casting cell value to property type. If this is true, invalid casts are silently skipped, otherwise any error will cause method to fail with exception.|
**Returns**: A list of objects from given worksheet



---
## ToExcel extensions
---
#### Method ToExcelExtensions.ToWorksheet<T>()

Generates a worksheet wrapper object from given list

|Name | Description |
|-----|------|
|T: | Type of object |
|rows: | List of objects |
|name: | Name of worksheet.|
|configureColumn: ||
|configureHeader: ||
|configureHeaderRow: ||
|configureCell: ||
**Returns**: 



---
#### Method ToExcelExtensions.NextWorksheet<T>()

 Starts new worksheet on same Excel package 

|Name | Description |
|-----|------|
|T: ||
|K: ||
|previousSheet: ||
|rows: ||
|name: ||
|configureColumn: ||
|configureHeader: ||
|configureHeaderRow: ||
|configureCell: ||
**Returns**: 



---
#### Method ToExcelExtensions.WithColumn<T>()

 Adds a column mapping. If no column mappings are specified all public properties will be used 

|Name | Description |
|-----|------|
|T: ||
|worksheet: ||
|map: ||
|columnHeader: ||
|configureColumn: ||
|configureHeader: ||
|configureCell: ||
**Returns**: 



---
#### Method ToExcelExtensions.WithTitle<T>()

 Adds a title row to the top of the sheet 

|Name | Description |
|-----|------|
|T: ||
|worksheet: ||
|title: ||
|configureTitle: ||
**Returns**: 



---
#### Method ToExcelExtensions.ToPackage<T>()

 Converts given list of objects to ExcelPackage 

|Name | Description |
|-----|------|
|T: ||
|rows: ||
**Returns**: 

#### Method ToExcelExtensions.ToPackage<T>()


|Name | Description |
|-----|------|
|T: ||
|lastWorksheet: ||
**Returns**: 

---
#### Method ToExcelExtensions.ToXlsx<T>()



|Name | Description |
|-----|------|
|T: ||
|rows: ||
**Returns**: A byte array to save the given list as Excel file

---
#### Method ToExcelExtensions.ToXlsx<T>()

|Name | Description |
|-----|------|
|T: ||
|lastWorksheet: ||
**Returns**: 

---
### **Examples**
---