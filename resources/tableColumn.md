# TableColumn
Represents a Column in a table. The TableColumn object is a member of the TableColumns collection.

## JSON representation 

JSON representation of a Table Column resource
<!-- { "blockType": "resource", "@odata.type": "TableColumn", 
		"optionalProperties": ["totalRange", "dataBodyRange"],	 
		"nullableProperties": [ "values"]
	 } 
-->
```json
{
  "id" : 999,
  "index" : 1,
  "name" : "String",
  "totalsCalculation" : "String",
  "values" : [[ "values" ]],

  "totalRange":  {"@odata.type": "Range"},
  "dataBodyRange":  {"@odata.type": "Range"}
}
```

## Properties

|Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
| `id`     |  Number | A unique key that identifies the Table Column within the Table   |        |
|`index`          |  Number | Index number of the TableColumn object within the TableColumns collection. **Zero Indexed**| ListColumn.Index|
|`name`           | String | String value that represents the name of the Table column..| ListColumn.Name|
|`totalsCalculation` |String | Constant value that determines the type of calculation in the Totals row of the list column. Possible values are: `Average`, `Count`, `CountNums`, `Max`, `Min`, `None`, `Sum`, `StdDev`, `Var`
|`values`         | Array (Primitive)  | Unformatted values of the table Column. |Collection of ListColumn.Range.Value2|


## Relationships
The TableColumn resource has the following relationships defined:

| Relationships    | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
| `totalRange`  | [Range](range.md) Object |Returns the Total row for a Column Object|ListColumn.Range|ListColumn.Total|
| `dataBodyRange`  | [Range](range.md) Object |Returns a Range object that is the size of the data portion of a column.|ListColumn.DataBodyRange|

## Methods

The complete list of methods for this resource is available in
the [API](../README.md) topic.