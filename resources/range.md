# Range
Range represents a cell, a row, a column, a selection of cells containing one or more contiguous blocks of cells. 


## JSON representation

JSON representation of a Range resource.
<!-- { "blockType": "resource", "@odata.type": "Range", 
	"optionalProperties": ["rows", "columns", "format", "areas", "values"],
  "nullableProperties": [ "values", "text", "numberFormat", "formula", "formulaLocal", "hasFormula" ]
	 } 
-->
```json
{
  "cellCount" : 99,
  "columnIndex" : 99,
  "rowIndex" : 99,
  "values" : [[ "String" , 99]],
  "text" : [ [ "Text" ]],
  "numberFormat" : [ [ "Number Format" ]],
  "formula" : [ ["Formula" ] ],
  "formulaLocal" : [ [ "Formula" ]],
  "address" : "Address",
  "addressLocal" : "Address Locale",

  "rows" :    [ {"@odata.type": "Range"} ],
  "columns" : [ {"@odata.type": "Range"} ],
  "format"  : { "@odata.type" : "Format" },
  "entireRow" : { @odata.type": "Range"}
  "entireColumn" : { @odata.type": "Range"}
}
```

## Properties
| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`cellCount`       | Number          |Number of cells in the range|Range.Count|
|`columnIndex`     | Number          |Returns the number of the first column in the first area in the specified range. This is adjusted to be zero indexed. Read-only|Range.Column|
|`rowIndex`        | Number          |Returns the number of the first row of the first area in the range. This is adjusted to be zero indexed. Read-only|Range.Row|
|`rowcount`        | Number          |Returns the total number of columns in the Range selected. Read-only |Range.Column|
|`columnCount`    | Number           |Returns the number of the first row of the first area in the range. This is adjusted to be zero indexed. Read-only|Range.Row|
|`values`          |Array [][]]|Unformatted values of the specified range|Range.Value2|
|`text`            |Array [][]|Text values of the specified range. The Text value will not depend on the cell width. The # sign substitution that happens in Excel client will not affect the value returned by the API. |Range.Text|
|`numberFormat`    |Array [][]|Value that represents the format code for the object|Range.NumberFormat
|`formula`         |Array [][]|Represents the object's formula in A1 style notation|Range.formula|
|`formulaLocal`    |Array [][]|Formula for the object, in the language of the user in A1 style notation|Range.FormulaLocal|
|`address`         |String         |Returns a String value that represents the range reference in A1 Style. **Address value will contain the Sheet reference (e.g., `Sheet1!A1:B4`)**|Range.Address|
|`addressLocal`    |String         |Returns the range reference for the specified range in the language of the user in A1 Style. **Address value will contain the Sheet reference (e.g., `Sheet1!A1:B4`)**|Range.AddresLocal|

## Relationships
Range resource has the following relationships defined:

| Relationships    | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`entireRow`            |[Range](range.md) collection| Range that represents the entire row(s) of the current range|Range.entireRow|
|`entireColumn`            |[Range](range.md) collection| Range that represents the entire column(s) of the current range|Range.entireRow|
|`rows`            |[Range](range.md) collection| Collection rows that make up the Range |Range.Rows|
|`columns`         |[Range](range.md) collection| Collection rows that make up the Range |Range.Columns|
|`format`          |[Format](format.md) Object  |Format object contains Range's Font, Background, Borders, Alignment, Style, etc. settings |Custom Object|


## Methods

The complete list of methods for this resource is available in
the [API](../README.md) topic.