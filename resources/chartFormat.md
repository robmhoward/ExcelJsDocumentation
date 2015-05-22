# ChartFormat
Provides access to the Office Art formatting for chart elements.


## JSON representation

JSON representation of a ChartFormat resource.
<!-- { "blockType": "resource", "@odata.type": "ChartFormat", 
	"optionalProperties": ["fill", "line", "font"]
	 } 
-->
```json
{

  "fill" :  {"@odata.type": "Fill"} ,
  "line" :  {"@odata.type": "Line"},
  "font" :  {"@odata.type": "Font"}

}
```

## Properties

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|



## Relationships
The Chart resource has the following relationships defined:

| Relationships    | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
| `fill`          |[Fill](fill.md) Object | Represents the fill format of an object, which includes background formating information. 
| `line`          |[Line](line.md) Object | Represents line and arrowhead formatting.
| `font`          |[Font](font.md) Object | Represents line and arrowhead formatting.
     

## Methods

The complete list of methods for this resource is available in
the [API](../README.md) topic.