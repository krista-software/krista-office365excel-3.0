# Krista Office 365 Excel Extension

## Overview

The **Krista Office 365 Excel Extension** enables seamless integration with Microsoft Office 365 Excel, allowing you to automate spreadsheet operations, manage workbooks, sheets, and tables, and perform data manipulation tasks directly from your Krista workflows.

This extension leverages the Microsoft Graph API to provide comprehensive Excel automation capabilities for both shared drives and personal OneDrive files.

## Key Features

- ✅ **Workbook Management**: Access and manage Excel files from OneDrive and shared drives
- ✅ **Sheet Operations**: Create, delete, read, and modify worksheets
- ✅ **Table Operations**: Perform CRUD operations on Excel tables
- ✅ **Data Manipulation**: Add, update, delete rows and columns with flexible input options
- ✅ **Advanced Functions**: Execute formulas, VLOOKUP operations, and range calculations
- ✅ **Batch Operations**: Asynchronous processing for large datasets
- ✅ **File Conversion**: Convert Excel files to PDF format
- ✅ **Flexible Authentication**: Support for both service account and user account authentication
- ✅ **Retry Mechanism**: Built-in retry logic for handling transient failures

## Quick Start Guide

### 1. Prerequisites
- Microsoft Azure account with administrative access
- Office 365 subscription with Excel Online
- Access to create Microsoft Entra ID applications

### 2. Setup Steps
1. [Create a Microsoft Azure Application](pages/CreatingMicrosoftApp.md)
2. [Configure the Extension](pages/ExtensionConfiguration.md)
3. [Set Up Authentication](pages/Authentication.md)

### 3. Start Using
Once configured, you can use any of the [catalog requests](#catalog-requests) in your Krista workflows.

## Documentation Structure

### Getting Started
- [Extension Configuration](pages/ExtensionConfiguration.md) - Complete setup and configuration guide
- [Authentication](pages/Authentication.md) - OAuth 2.0 authentication, token management, and security
- [Creating Microsoft App](pages/CreatingMicrosoftApp.md) - Step-by-step Microsoft Entra ID application setup

### Catalog Requests

#### 📊 Shared Drive Operations
- [Get All Shared Drives Information](pages/GetAllSharedDrivesInformation.md)
- [Get Shared Excel Files](pages/GetSharedExcelFiles.md)
- [Add Rows to Table](pages/AddRowsToTableShared.md)
- [Add Rows to Table (Async)](pages/AddRowsToTableAsync.md)
- [Get Add Rows To Table Result](pages/GetAddRowsToTableResult.md)
- [Add Row With List of Entities](pages/AddRowWithListOfEntities.md)
- [Add Rows to Sheet](pages/AddRowsToSheet.md)
- [Add Row with Index](pages/AddRowWithIndex.md)
- [Add Row with Index (Async)](pages/AddRowWithIndexAsync.md)
- [Get Add Row with Index Result](pages/GetAddRowWithIndexResult.md)
- [Get Rows for Range](pages/GetRowsForRangeShared.md)
- [Get Table Rows](pages/GetTableRows.md)
- [Get Column Values for Keys](pages/GetColumnValuesForKeysShared.md)
- [Get Column Values for Keys in Sheet](pages/GetColumnValuesForKeysInSheet.md)
- [Add Data to Provided Cell](pages/AddDataToProvidedCell.md)
- [Add Multiple Cell Data](pages/AddMultipleCellData.md)
- [Cell Data](pages/CellData.md)
- [Search](pages/Search.md)
- [Range Sum](pages/RangeSum.md)
- [Sheet VLOOKUP](pages/SheetVLOOKUP.md)
- [VLOOKUP in Table](pages/VLOOKUPInTable.md)
- [Columns Sum](pages/ColumnsSum.md)
- [Execute Formula for Shared File](pages/ExecuteFormulaForSharedFile.md)
- [Update Row with Index](pages/UpdateRowWithIndex.md)
- [Delete Row with Index](pages/DeleteRowWithIndexShared.md)
- [Delete Column with Index](pages/DeleteColumnWithIndexShared.md)
- [Get Complete Row Info](pages/GetCompleteRowInfo.md)
- [Convert to PDF](pages/ConvertToPDF.md)

#### 📄 Sheet Operations
- [Get Excel File Information](pages/GetExcelFileInformation.md)
- [Get All Excel Files For Name](pages/GetAllExcelFilesForName.md)
- [Get Sheets Information](pages/GetSheetsInformation.md)
- [Create Sheet](pages/CreateSheet.md)
- [Delete Sheet](pages/DeleteSheet.md)
- [Get Cell Value](pages/GetCellValue.md)
- [Add Data to Cell](pages/AddDataToCell.md)
- [Add Rows](pages/AddRows.md)
- [Get Rows For Range](pages/GetRowsForRange.md)
- [Get Rows For Range By Path](pages/GetRowsForRangeByPath.md)
- [Clear Sheet For Range](pages/ClearSheetForRange.md)
- [Find Value](pages/FindValue.md)
- [Vlookup](pages/Vlookup.md)
- [Execute Formula](pages/ExecuteFormula.md)
- [Get Column Values for Keys In Sheet](pages/GetColumnValuesForKeysInSheetArea.md)
- [Delete Row](pages/DeleteRow.md)
- [Delete Column](pages/DeleteColumn.md)
- [Update Row](pages/UpdateRow.md)
- [Get Row](pages/GetRow.md)
- [Add Row At Index](pages/AddRowAtIndex.md)
- [Add Row At Index (Async)](pages/AddRowAtIndexAsync.md)
- [Get Add Row At Index Result](pages/GetAddRowAtIndexResult.md)

#### 📋 Table Operations
- [Table Vlookup](pages/TableVlookup.md)
- [Add Rows](pages/AddRowsTable.md)
- [Add Rows Async](pages/AddRowsAsync.md)
- [Get Add Rows Result](pages/GetAddRowsResult.md)
- [Add Rows To Table](pages/AddRowsToTable.md)
- [Add Rows To Table Async](pages/AddRowsToTableAsyncTable.md)
- [Get Add Rows To Table Result](pages/GetAddRowsToTableResultTable.md)
- [Get Rows](pages/GetRows.md)
- [Summation](pages/Summation.md)
- [Delete Column](pages/DeleteColumnTable.md)
- [Create Table for Range](pages/CreateTableForRange.md)
- [Get Column Values for Keys](pages/GetColumnValuesForKeys.md)
- [Add Rows To Table In Batch Async](pages/AddRowsToTableInBatchAsync.md)
- [Add Rows Async To Table Progress Update](pages/AddRowsAsyncToTableProgressUpdate.md)
- [Add Rows Async To Table Result Available](pages/AddRowsAsyncToTableResultAvailable.md)

### Additional Resources
- [Troubleshooting Guide](pages/Troubleshooting.md) - Common issues and solutions
- [Release Notes](pages/releaseNotes.md) - Version history and updates

## Support & Resources

- **Extension Version**: 3.0.4
- **Krista Service APIs**: 1.0.121
- **Global Catalog Version**: GC-2026.06.1
- **Microsoft Graph API**: v1.0
- **Required Scopes**: Files.ReadWrite, Files.ReadWrite.All, Sites.Read.All

## Error Handling & Troubleshooting

The Office365Excel extension provides user-friendly error messages to help you quickly identify and resolve issues:

### Common Error Types

#### Authentication Errors
- **Session Expired**: Re-authorize the connection through the extension configuration
- **Permission Issues**: Contact your system administrator to verify application permissions
- **Invalid Credentials**: Check that your client ID and secret are correct and up-to-date

#### File Access Errors
- **File Not Found**: Verify the file ID is correct and the file exists
- **Permission Denied**: Ensure you have read/write access to the Excel file
- **Invalid File Format**: Confirm the file has a valid Excel extension (.xlsx, .xlsm, .xls, .xlsb)

#### Data Validation Errors
- **Invalid Cell Reference**: Use proper Excel cell format (e.g., A1, B2, AA10)
- **Invalid Range**: Use valid Excel range formats (A1:B10, A:A, 1:1)
- **Sheet Not Found**: Check sheet name spelling and use 'Get Sheets Information' to verify
- **Table/Column Not Found**: Verify table and column names are case-sensitive and exist

#### Formula Errors
- **Invalid Formula**: Ensure formulas start with '=' and use valid Excel syntax
- **Formula Too Long**: Keep formulas under 8,192 characters
- **Reference Errors**: Verify all cell references in formulas exist

### Getting Help

- Check the [Troubleshooting Guide](pages/Troubleshooting.md) for common issues and solutions
- Review the [Authentication Guide](pages/Authentication.md) for authentication issues
- Check the [Extension Configuration](pages/ExtensionConfiguration.md) for setup problems
- Consult individual catalog request pages for specific operation details
- Review [Release Notes](pages/releaseNotes.md) for known issues and bug fixes
- Contact your system administrator for permission-related issues
