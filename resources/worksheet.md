# Worksheet
The Worksheet object is a member of the Worksheets collection. The Worksheets collection contains all the Worksheet objects in a workbook.

## JSON representation 

JSON representation of a Worksheet resource

<!-- { "blockType": "resource", "@odata.type": "Worksheet", 
		"optionalProperties": ["usedRange", "tables"]
	 } 
-->
```json
{
  "index" : 1,
  "name" : "String",

  "usedRange":  {"@odata.type": "Range"},
  "tables":  [{"@odata.type": "Table"}],
  "charts":  [{"@odata.type": "Chart"}]
  " entireWorksheetRange":  [{"@odata.type": "Range"}]

}
```

## Properties

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
| `index`|  Number |The zero-based index of the worksheet within the workbook|Worksheet.Index|
| `name` | String|The user-visible name of the worksheet|Worksheet.Name |


## Relationships
The Worksheet resource has the following relationships defined:

| Relationship     | Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
| `tables`        | [Table](table.md) collection |Collection of tables in the worksheet|Worksheet.ListObjects  |          
| `charts`        | [Chart](chart.md) collection |Collection of charts in the worksheet|Worksheet.ChartObject  |       

## Methods

The Worksheet resource has the following methods defined:

| Method     | Return Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
| `activate()`     | void     |Activates the worksheet |   |
| `deleteObject()`     | void     |Deletes the worksheet ||
| `getCell(row: number, column: number)`        | [Range](range.md) object |Returns a range containing the single cell specified by the zero-indexed row and column numbers| |          
| `getEntireWorksheetRange()`        | [Range](range.md) object |Returns the range containing all cells in the worksheet| |
| `getRange(address: string)`        | [Range](range.md) object |Returns the range specified by the address| |
| `getUsedRange()`        | [Range](range.md) object |Returns the used range of the worksheet| |  