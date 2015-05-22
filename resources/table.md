# Table
Represents collection of organized cells designed to make management of the data easy.

## JSON representation  

JSON representation of a Table resource.
<!-- { "blockType": "resource", "@odata.type": "Table", 
		"optionalProperties": ["tableRows", "tableColumns", "headerRowRange", "dataBodyRange", "totalsRowRange", "range"],	 
	 } 
-->
```json
{
  "id" : 999,
  "name" : "String",
  "displayName" : "String",
  "showTotals" : false,
  "tableStyle" : "String",

  "tableRows":      [ {"@odata.type": "TableRow"} ],
  "tableColumns":   [ {"@odata.type": "TableColumn"} ],
  "headerRowRange": { "@odata.type" : "Range" },
  "dataBodyRange":  {"@odata.type": "Range"},
  "totalsRowRange": {"@odata.type": "Range"},
  "range" : 	 	{"@odata.type": "Range"} 
}
```

## Properties

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
| `id`  |  Number | A unique key that identifies the Table object in a Workbook. Note: If the table gets deleted, Excel Workbook could re-use the id value for another Table.   |        |
| `name`       | String| String value that represents the name of the Table object   | ListObject.Name       |
| `displayName`| String| Display name for the specified Table                        | ListObject.DisplayName|
| `showTotals` | Boolean| Boolean to indicate whether the Total row is visible. This value can be set to show or remove the total row| ListObject.ShowTotals|
| `tableStyle` | String | Constant that represents the Table style. Possible values include: `Light1` thru `Light21`, `Medium1` thru `Medium28`, `StyleDark1` thru `StyleDark11`|ListObject.TableStyle|


## Relationships
The Table resource has the following relationships defined:

| Relationships    | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
| `tableRows`      | TableRow collection         |Collection of Table rows |ListObject.ListRows      |
| `tableColumns`   | TableColumn collection      |Collection of Table columns |ListObject.TableColumns  |          
| `headerRowRange` | [Range](range.md) Object    |Header row's Range object |ListObject.HeaderRowRange|        
| `dataBodyRange`  | [Range](range.md) Object    |Table body's Range object |ListObject.DataBodyRange |
| `totalsRowRange` | [Range](range.md) Object    |Total row's range object |ListObject.TotalsRowRange|
| `range`          | [Range](range.md) Object    |Table's Range object |ListObject.Range         |

## Methods

The complete list of methods for this resource is available in
the [API](../README.md) topic.