# TableRow
Represents a row in a table. The TableRow object is a member of the TableRows collection.



## Properties

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
| `index`          |  Number | Index number of the TableRow object within the TableRows collection. **Zero Indexed**| ListRow.Index|
| `values`         | Array (Primitive)  | Unformatted values of the table row. |Collection of ListRow.Range.Value2|


## Relationships
The TableRow has the following relationships defined:

| Relationships    | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
| `range`  | [Range](range.md) Object |Returns a Range object associated with the Table Row.|ListRow.Range|

## Methods

None