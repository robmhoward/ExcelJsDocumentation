# Axis
Represents the collection in a Chart including categoryAxis, valueAxis and seriesAxis.


## JSON representation

JSON representation of a Range resource.
<!-- { "blockType": "resource", "@odata.type": "ChartAxes", 
	"optionalProperties": ["categoryAxis", "valueAxis", "seriesAxis" ]
	 } 
-->
```json
{
  "categoryAxis" :    {"@odata.type": "ChartAxis"} ,
  "valueAxis"  : { "@odata.type" : "ChartAxis"},
  "seriesAxis" : {"@odata.type": "ChartAxis"}  

}
```

## Properties

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|



## Relationships
The Chart resource has the following relationships defined:

| Relationships    | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
| `categoryAxis`          |[ChartAxis](chartAxis.md) Object | Represents the category usually horizontal axis in a chart. | 
| `valueAxis` | [ChartAxis](chartAxis.md) Object   | Represents the value axis in a chart.  | |
| `seriesAxis` | [ChartAxis](chartAxis.md) Object   |Represents the series axis in a 3D chart. | |
     

## Methods

The complete list of methods for this resource is available in
the [API](../README.md) topic.