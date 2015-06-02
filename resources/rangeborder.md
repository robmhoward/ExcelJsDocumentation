
# Range Border

Represents the border of an object. 

## Properties
| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|sideIndex| String | Constant value that indicates the specific side of the border. Possible values are:  `DiagonalDown`, `DiagonalUp`, `EdgeBottom`, `EdgeLeft`, `EdgeRight`, `EdgeTop`, `InsideHorizontal`, `InsideVertical`.| String value derived from xlBordersIndex constants|
|lineStyle| String | One of the constants of LineStyle specifying the line style for the border. Options are: `Continuous`: Continuous line, `Dash`: Dashed line, `DashDot`: Alternating dashes and dots, `DashDotDot`: Dash followed by two dots, `Dot`: Dotted line, `Double`: Double line, `LineStyleNone`: No line, `SlantDashDot`: Slanted dashes.|Border.LineStyle|
|weight| String | BorderWeight value that specifies the weight of the border around a range. Options are: `Hairline`: Hairline (thinnest border), `Medium`: Medium, `Thick`: Thick (widest border), `Thin`: Thin.|Border.Weight|
|color| String | HTML color code representing the color of the border line|Border.Color's representation in HTML color code.|


`sideindex` specifies type of border to be retrieved/set as part of Borders collection. 

|SideIndex|Description|
|:--------|:----------|
|`InsideHorizontal`|Horizontal borders for all cells in the range except borders on the outside of the range.|
|`InsideVertical`  |Vertical borders for all the cells in the range except borders on the outside of the range.|
|`DiagonalDown`    |Border running from the upper left-hand corner to the lower right of each cell in the range.|
|`DiagonalUp`      |Border running from the lower left-hand corner to the upper right of each cell in the range.|
|`EdgeBottom`      |Border at the bottom of the range.|
|`EdgeLeft`        |Border at the left-hand edge of the range.|
|`EdgeRight`       |Border at the right-hand edge of the range.|
|`EdgeTop`         |Border at the top of the range.|


## Relationships
None

## Methods
None


