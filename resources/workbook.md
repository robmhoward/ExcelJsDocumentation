# Workbook
Workbook is the top level object which contains related workbook objects such as worksheets, tables, ranges, etc. It can be used to list related references. 

## JSON representation 

JSON representation of a Workbook resource

<!-- { "blockType": "resource", "@odata.type": "Workbook", 
		"optionalProperties": ["charts", "names", "tables", "worksheets"]
	 } 
-->
```json
{
  "charts":  [{"@odata.type": "Chart"}],
  "names":  [{"@odata.type": "NamedItem"}],
  "tables":  [{"@odata.type": "Table"}],
  "worksheets":  [{"@odata.type": "Worksheet"}],
  "activeWorksheet": {"@odata.type": "Worksheet"},
  "application": {"@odata.type": "Application"} 
}
```

## Properties

None

## Relationships
The Workbook resource has the following relationships defined:

| Relationship    | Type    |Description|Notes |
|:----------------|:--------|:----------|:-----|
| `charts`        | [Chart](chart.md) collection        | Collection of Charts associated with the workbook        |Workbook.Charts     |
| `names`         | [NamedItem](nameditem.md) collection| Collection of Named Ranges associated with the workbook  |Workbook.Names      |
| `tables`        | [Table](table.md) collection        | Collection of Tables associated with the workbook        |Workbook.ListObjects|
| `worksheets`    | [Worksheet](worksheet.md) collection| Collection of Worksheets associated with the workbook    |Workbook.Worksheets |
| `activeWorksheet`    | [Worksheet](worksheet.md)| Returns an object that represents the active sheet in the workbook. Returns `null` if no worksheet is active or a chart-sheet is active. |
| `application`    | [Application](application.md)| Returns an object that represents epresents the Excel application which is managing the workbook. |

## Methods

The complete list of methods for this resource is available in
the [API](../README.md) topic.