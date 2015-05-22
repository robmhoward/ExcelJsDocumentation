# Gridlines
Represents major or minor gridlines on a chart axis.


## JSON representation

JSON representation of a Range resource.
<!-- { "blockType": "resource", "@odata.type": "ChartGridlines", 
	"optionalProperties": "lineFormat"
	 } 
	 
-->
```json
{
  "visible" : true

}
```

## Properties

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|visible| Boolean | True if the axis has gridlines. |Axis.HasMajorGridlines|

## Relationships
The ChartGridlines resource has the following relationships defined:

| Relationships    | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
| `lineFormat`          |[ChartLineFormat](chartLineFormat.md) Object | Represents line and arrowhead formatting.

          

## Methods

The complete list of methods for this resource is available in
the [API](../README.md) topic.