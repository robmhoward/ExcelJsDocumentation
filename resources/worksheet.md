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
| `index`|  Number | Returns a integer value that represents the index number of the object within the collection of similar objects. **Zero Indexed**|Worksheet.Index|
| `name` | String| A String value that represents a Worksheet object |Worksheet.Name |


## Relationships
The Worksheet resource has the following relationships defined:

| Relationship     | Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
| `usedRange`     | [Range](range.md) object     | used-range of the Worksheet                       |Worksheet.UsedRange    |
| ` entireWorksheetRange`     | [Range](range.md) object     | used-range of the Worksheet                       ||
| `tables`        | [Table](table.md) collection | Collection of Tables associated with the Worksheet|Worksheet.ListObjects  |          
| `charts`        | [Chart](chart.md) collection | Collection of Charts associated with the Worksheet|Worksheet.ChartObject  |       

## Methods

The complete list of methods for this resource is available in
the [API](../README.md) topic.