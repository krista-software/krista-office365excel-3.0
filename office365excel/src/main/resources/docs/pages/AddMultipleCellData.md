# Add Multiple Cell Data

## Overview

Updates multiple cells in a worksheet with different values in a single operation.

## Request Details

- **Area**: Shared
- **Type**: CHANGE_SYSTEM
- **Retry Support**: ✅ Yes

## Input Parameters

| Parameter Name | Type | Required | Description | Example |
|----------------|------|----------|-------------|---------|
| Drive Id | Text | Yes | Drive ID | "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8" |
| File Id | Text | Yes | File ID | "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROR" |
| Sheet Name | Text | Yes | Sheet name | "Sheet1" |
| Cell ID Data | List&lt;{Cell Id: Text, Data: Text}&gt; | Yes | List of cell IDs and values | [{"Cell Id": "A1", "Data": "Hello"}] |

### Cell ID Data Structure

Each element in the list must contain:
- **Cell Id**: Excel cell reference (e.g., "A1", "B5")
- **Data**: Value to write to the cell

## Output Parameters

| Parameter Name | Type | Description |
|----------------|------|-------------|
| Cell Data | List&lt;Composite&gt; | Updated cell data |

## Usage Examples

### Example 1: Update Multiple Cells

**Input**:
```
Drive Id: "b!tVhOAm0EY0WK0z3E4xcAH5Ufuu38F8"
File Id: "01VZSNCFS7ZQ7OXAN22BDL43PIPMRALROR"
Sheet Name: "Sheet1"
Cell ID Data: [
  {"Cell Id": "A1", "Data": "Name"},
  {"Cell Id": "B1", "Data": "Age"},
  {"Cell Id": "A2", "Data": "John"},
  {"Cell Id": "B2", "Data": "30"}
]
```

**Output**:
```json
{
  "Cell Data": [
    {"Cell Id": "A1", "Value": "Name"},
    {"Cell Id": "B1", "Value": "Age"},
    {"Cell Id": "A2", "Value": "John"},
    {"Cell Id": "B2", "Value": "30"}
  ]
}
```

## Business Rules

1. **Batch Update**: Updates multiple cells in one operation
2. **Overwrites**: Replaces existing cell values
3. **Any Position**: Can update cells at any position
4. **Formula Support**: Can write formulas (start with =)

## Limitations

1. **Batch Size**: Large batches may hit API limits
2. **Shared Drives Only**: For shared drive files only

## Best Practices

### 1. Validate Cell IDs
Ensure all cell IDs are valid Excel references.

### 2. Use for Scattered Updates
Use this when updating non-contiguous cells.

### 3. Batch Related Updates
Group related cell updates in one operation.

## Related Catalog Requests

- [Add Multiple Cell Data Async](pages/AddMultipleCellDataAsync.md) - Async version
- [Add Data to Provided Cell](pages/AddDataToProvidedCell.md) - Single cell update
- [Get Cell Value](pages/GetCellValue.md) - Read cell values

## Technical Implementation

### Helper Class
- **Class**: SharedOperationsFacade
- **Package**: app.krista.extensions.essentials.spreadsheets.office365excel.catalog.service

### Microsoft Graph API Endpoint
- **Endpoint**: `PATCH /drives/{driveId}/items/{fileId}/workbook/worksheets/{sheetName}/range(address='{cellId}')`
- **Permissions Required**: Files.ReadWrite.All

## See Also

- [Extension Configuration](pages/ExtensionConfiguration.md)
- [Authentication](pages/Authentication.md)
- [Add Multiple Cell Data Async](pages/AddMultipleCellDataAsync.md)

