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

  "format"  : { "@odata.type" : "Format" },
  "worksheet"  : { "@odata.type" : "Worksheet" }
}
```

## Properties
| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`cellCount`       | Number          |Number of cells in the range|Range.Count|
|`columnIndex`     | Number          |Zero-based index of the first column in the first area of the range. Read-only|Range.Column|
|`rowIndex`        | Number          |Zero-based index of the first row in the first area of the range. Read-only|Range.Row|
|`rowCount`        | Number          |Total number of rows in the range. Read-only |Range.Column|
|`columnCount`    | Number           |Total number of columns in the range. Read-only |Range.Row|
|`values`          |Array [][]]|Array of arrays representing the unformatted values of the cells in the range|Range.Value2|
|`text`            |Array [][]|Array of arrays representing the formatted text values of the cells in the range. The text value will not depend on the cell width. The # sign substitution that happens in Excel client will not affect the value returned by the API.|Range.Text|
|`numberFormat`    |Array [][]|Array of arrays representing the format code for each of the cells in the range|Range.NumberFormat
|`formulas`         |Array [][]|Array of arrays representing the formulas in the range's cells using A1 style notation. Setting to a single value applies to all cells.|Range.formula|
|`formulasLocal`    |Array [][]|Array of arrays representing the formulas in the range's cells using A1 style notation in the user's language. Setting to a single value applies to all cells.|Range.FormulaLocal|
|`address`         |String         |Returns a String value that represents the range reference in A1 Style. **Address value will contain the Sheet reference (e.g., `Sheet1!A1:B4`)**|Range.Address|
|`addressLocal`    |String         |Returns the range reference for the specified range in the language of the user in A1 Style. **Address value will contain the Sheet reference (e.g., `Sheet1!A1:B4`)**|Range.AddresLocal|

## Relationships
Range resource has the following relationships defined:

| Relationships    | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`entireRow`            |[Range](range.md) collection| Range that represents the entire row(s) containing the current range|Range.entireRow|
|`entireColumn`            |[Range](range.md) collection| Range that represents the entire column(s) containing the current range|Range.entireRow|
|`format`          |[Format](format.md) Object  |Format object contains Range's Font, Background, Borders, Alignment, Style, etc. settings |Custom Object|
|`worksheet`          |[Worksheet](worksheet.md) Object  | The worksheet containing the current range |


## Methods

The complete list of methods for this resource is available in
the [API](../README.md) topic.