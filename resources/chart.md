# Chart
Represents a chart object in a workbook.

## JSON representation

JSON representation of a Chart resource.

<!-- { "blockType": "resource", "@odata.type": "Chart","optionalProperties": ["title","series","axes", "dataLabels", "legend",  "fillFormat", "lineFormat", "font" ]
} 
-->
```json
{
  "name": "Chart1",
  "height" : 99,
  "width" : 99,
  "top" : 99,
  "left" : 99,

  "title" :     {"@odata.type": "ChartTitle"} ,
  "series" : {"@odata.type": "ChartSeries"} ,
  "axes" :     {"@odata.type": "ChartAxes"} ,
  "dataLabels"  : { "@odata.type" : "ChartDataLabels" },
  "legend" :    {"@odata.type": "ChartLegend"},
  "fillFormat" :    {"@odata.type": "ChartFillFormat"},
  "lineformat" :    {"@odata.type": "ChartLineFormat"},
  "font" :    {"@odata.type": "ChartFont"}

  }
```

## Properties

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
| `name`  | String | A String value that represents the name of a Chart object.   | Chart.Name      |
| `height`| Double | A Double value that represents the height, in points, of the chart object. | ChartArea.Height|
| `width` | Double | A Double value that represents the width, in points, of the chart object. | ChartArea.Width |
| `top` | Double |a Double value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).| ChartArea.Top |
| `left` | Double | a Double value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).| ChartArea.Left |


## Relationships
The Chart resource has the following relationships defined:

| Relationships    | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
| `title`          |[ChartTitle](chartTitle.md) Object | Returns a ChartTitle object that represents the title of the specified chart, including the text, visibility, position and formating of the title.
| `series`          |[ChartSeries](chartseries.md) Object |Represents a series in a chart.
| `axes`          |[ChartAxes](axes.md) Object |Represents a collection of Axes in the Chart.
| `dataLabels`          |[ChartDataLabels](chartDataLabels.md) Object | Represents the datalabels on the chart.
| `legend`          |[ChartLegend](chartLegend.md) Object |Returns a Legend object that represents the legend for the chart. 
| `fillFormat`          |[ChartFillFormat](chartFillFormat.md) Object | Represents the fill format of an object, which includes interior/background formating information. 
| `lineFormat`          |[ChartLineFormat](chartLineFormat.md) Object | Represents line and arrowhead formatting.
| `font`          |[ChartFont](chartFont.md) Object | Represents the font attributes (font name, font size, color, and so on) for an object. 



     

## Methods

The complete list of methods for this resource is available in
the [API](../README.md) topic.