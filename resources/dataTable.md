# DataTable
Represents a chart data table.


## JSON representation

JSON representation of a Range resource.
<!-- { "blockType": "resource", "@odata.type": "DataTable", 
	"optionalProperties": ["format"]
	 } 
-->
```json
{
  "visible" : true,
  "showLegendKey" : true
}
```

## Properties

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
| `visible` | Boolean |True if the data table is visible.  | |
| `showLegendKey` | Boolean |True if the data label legend key is visible.  | DataTable.ShowLegendKey |




## Relationships
The Chart resource has the following relationships defined:

| Relationships    | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
| `format`          |[ChartFormat](chartFormat.md) Object | Provides access to the Office Art formatting for chart elements.


## Methods

The complete list of methods for this resource is available in
the [API](../README.md) topic.