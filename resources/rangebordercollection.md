# Range Border Collection 

Represents the border objects that make up Range border. 


## [Properties](#get-border-collection)

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`count`| Number   | Number of objects in the collection.|range.borders.count|
|`items`| Object[] | A collection of all the border objects that are part of the workbook|ListObjects |

## Relationships

None

## Methods

The border collection has the following methods defined:

| Method     | Return Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|[getItem(name: string)](#getitemname-string)| [border](rangeborder.md)      |Retrieve a border object using its name||
|[getItemAt(index: number)](#getitematindex-number)| [border](rangeborder.md)     |Retrieve a border based on its position in the items[] array.||


## API Specification 

### Get border Collection

Get properties of the border collection. 

#### Syntax
```js
rangeObject.borders.property;
```

#### Properties

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`count`| Number   | Number of objects in the collection.|borders.count|
|`items`| Object[] | A collection of all the border objects that are part of the workbook|[borders.item] |


#### Returns

[border](rangeborder.md) collection. 

#### Examples

```js
var sheetName = "Sheet1";
var rangeAddress = "D5:F8";
var ctx = new Excel.ExcelClientContext();
var worksheet = ctx.workbook.worksheets.getItem(sheetName);
var range = worksheet.getRange(rangeAddress);
var borders = range.format.borders;
ctx.load(borders);
ctx.executeAsync().then(function () {
	for (var i = 0; i < borders.items.length; i++)
	{
		Console.log(borders.items[i].sideIndex);
	}
});
```

##### Getting the number of borders

```js
var sheetName = "Sheet1";
var rangeAddress = "D5:F8";
var ctx = new Excel.ExcelClientContext();
var worksheet = ctx.workbook.worksheets.getItem(sheetName);
var range = worksheet.getRange(rangeAddress);
var borders = range.format.borders;
ctx.load(borders);
ctx.executeAsync().then(function () {
	Console.log(borders.count);
});
```
[Back](#properties)

### Set Range Border 

Sets border to a range and sets the Color, LineStyle, and Weight properties for the new border.

#### Syntax
```js
rangeObject.format.borders(sideIndex).property = value;
```
Where, property is one of the following Range's border properties that can be set. 

#### Properties

Property       | Type   | Description
--------------- | ------ | ------------
`lineStyle`| String | One of the constants of LineStyle specifying the line style for the border. Options are: `Continuous`: Continuous line, `Dash`: Dashed line, `DashDot`: Alternating dashes and dots, `DashDotDot`: Dash followed by two dots, `Dot`: Dotted line, `Double`: Double line, `None`: No line, `SlantDashDot`: Slanted dashes.|Border.LineStyle
`weight`| String | BorderWeight value that specifies the weight of the border around a range. Options are: `Hairline`: Hairline (thinnest border), `Medium`: Medium, `Thick`: Thick (widest border), `Thin`: Thin.|Border.Weight
`color`| String | HTML color code representing the color of the border line|Border.Color's representation in HTML color code.


**sideIndex values:**

`sideIndex` values | Type  | Description
--------------- | ------ | ------------
`DiagonalDown`  |String | Border running from the upper left-hand corner to the lower right of each cell in the range. 
`DiagonalUp`    |String |Border running from the lower left-hand corner to the upper right of each cell in the range.
`EdgeBottom`    |String |Border at the bottom of the range.
`EdgeLeft`      |String |Border at the left-hand edge of the range.
`EdgeRight`     |String |Border at the right-hand edge of the range.
`EdgeTop`       |String |Border at the top of the range.
`InsideHorizontal` |String|Horizontal borders for all cells in the range except borders on the outside of the range.
`InsideVertical`|String |Vertical borders for all the cells in the range except borders on the outside of the range.

#### Example
The example below adds grid border around the range.

```js
var sheetName = "Sheet1";
var rangeAddress = "F:G";
var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
range.format.borders('InsideHorizontal').lineStyle = 'Continuous';
range.format.borders('InsideVertical').lineStyle = 'Continuous';
range.format.borders('EdgeBottom').lineStyle = 'Continuous';
range.format.borders('EdgeLeft').lineStyle = 'Continuous';
range.format.borders('EdgeRight').lineStyle = 'Continuous';
range.format.borders('EdgeTop').lineStyle = 'Continuous';
ctx.executeAsync().then();
```
[Back](#properties)

### add(name: string)

Add a new border to the workbook. The border will be added at the end of existing borders.

#### Syntax
```js
bordersCollection.add(name);
```

#### Parameters

Parameter       | Type   | Description
--------------- | ------ | ------------
`name`  | String| Optional. String value representing the name of the sheet to be added. If not specified, Excel determines the name of the new border being added. 

#### Returns
[border](rangeborder.md) object.

#### Examples

```js
var wSheetName = 'Sample Name';
var ctx = new Excel.ExcelClientContext();
var border = ctx.workbook.borders.add(wSheetName);
ctx.load(border);
ctx.executeAsync().then(function () {
	Console.log(border.name);
});
```
[Back](#methods)

### getItem(name: string)

Get border object properties based on name.

#### Syntax
```js
bordersCollection.getItem(name);
```

#### Parameters

Parameter       | Type  | Description
--------------- | ------ | ------------
 `name`| String | Required. border name. 

#### Returns

[border](rangeborder.md) object.

#### Examples
```js
var ctx = new Excel.ExcelClientContext();
var borderName = 'border1';
var border = ctx.workbook.borders.getItem(borderName);
ctx.executeAsync().then(function () {
		Console.log(border.index);
});
```
[Back](#methods)


### getItemAt(index: number)

Get border object properties based on its position in the items[] array. 

#### Syntax
```js
bordersCollection.getItemAt(index);
```

#### Parameters

Parameter       | Type  | Description
--------------- | ------ | ------------
 `index`| Number | Required. Index or position in the items[]. Zero indexed.

#### Returns

[border](rangeborder.md) object.

#### Examples
```js
var ctx = new Excel.ExcelClientContext();
var border = ctx.workbook.borders.getItemAt(0);
ctx.executeAsync().then(function () {
		Console.log(border.name);
});
```
[Back](#methods)
