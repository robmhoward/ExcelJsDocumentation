# Font

This object represents the font attributes (font name, font size, color, and so on) for an object. 

## JSON representation

JSON representation of a Font resource.

<!-- { "blockType": "resource", "@odata.type": "ChartFont" } -->
```json
{
  "name" : "Ariel", 
  "size" : 14,
  "color" : "FF0000",
  "italic" : false,
  "bold" : false,
  "underline" : false 
}
```


## Properties
| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`name`|String|Font name (e.g., "Calibri")|Range.Font.Name|
|`size`|Double|Size of the font (e.g., 11)|Range.Font.Size|
|`color`|String|HTML color code representation of the text color. HTML color codes are strings that represents hexadecimal triplets of red, green, and blue values (#RRGGBB). e.g., `#FF0000` represents Red. ('255' red, '0' green, and '0' blue) |Conversion from Range.Font.Color value to html color string|
|`italic`|Boolean|Represents the bold status of italic. true if the font style is italic|Range.Font.Italic|
|`bold`|Boolean|Represents the bold status of font. true if the font is bold. |Range.Font.Bold|
|`underline`|Boolean|Type of underline applied to the font. Can be one of the following constants. Possible Values: `None`, `Single`, `Double`, `SingleAccounting`, `DoubleAccounting`|Range.Font.Underline|

## Relationships
None

## Methods

The complete list of methods for this resource is available in
the [API](../README.md) topic.
