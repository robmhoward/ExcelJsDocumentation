# ChartFillFormat
Represents the interior of an object, which includes background formating information. 

## Properties
| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|

## Programming notes about `color` property: 

A `color` hex code is a way of specifying color using hexadecimal values. The code itself is a hex triplet, which represents three separate values that specify the levels of the component colors. The code starts with a pound sign (#) and is followed by six hex values or three hex value pairs (for example, #AFD645). 

Of the 6 Hex values, first two characters represent the values 0 through 255 for red in hex; the middle two for green and the last two for blue (#RRGGBB). For example, FF is equal to 255. Therefore, the purest white obtainable is the highest intensity of red, green and blue, which is #FFFFFF (red=255, green=255 and blue=255). Black is the lack of all RGB (#0000000).

When `color` value is updated, the input value needs to follow the appropriate formatting as mentioned above. The Alpha characters of the hex color code can be lower or upper case. 

Alternatively,  `#` sign followed by 3 character color code (e.g., #F00) could be used to set the color. Note that the return color values are always coded as `#` followed by 6 character color code. 

## Relationships
None

## Methods

The complete list of methods for this resource is available in
the [API](../README.md) topic.


