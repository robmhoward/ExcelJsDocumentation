# Application

Represents the Excel application which is managing the workbook. 

## [Properties](#get-application)

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
| `calculationMode`        | String      | Specifies the Calculation mode of the workbook. Possible values are `Automatic`: Excel controls recalculation, `Manual`: Calculation is done when the user requests it, or `Semiautomatic`: Excel controls recalculation but ignores changes in tables.         |Workbook.Application.Calculation|


## Relationships
None

## Methods
The Application has the following methods defined:

| Method     | Return Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|[calculate()](#calculate)| [Worksheet](worksheet.md) object |Perform calculation on the workbook or application.| |

### Get Application

Get properties of workbook's application object. 

```js
workbookObject..application;
```
#### Returns

[Application](application.md) object.

#### Examples

```js
var ctx = new Excel.ExcelClientContext();
var application = ctx.workbook.application;
ctx.load(application);
ctx.executeAsync().then(function() {
	Console.log(application.calculationMode);
});

```
[Back](#properties)

### calculate()

Performs calculation on the workbook or application. 

#### Syntax
```js
applicationObject.calculate(calculationType)
```
#### Parameters

Parameter       | Type  | Description
--------------- | ------ | ------------
`calculationType` | String | Optional. Available Options are: `ReCalculate`: does normal calculation, `Full`: forces a full calculation of the data, `FullRebuild`: forces a full calculation of the data and rebuilds the dependencies (this is similar to re-entering all formulas). Note: if request body is not provided then calculation of the type `ReCalculation` is performed.

#### Returns

Nothing

#### Examples 

```js
var ctx = new Excel.ExcelClientContext();
ctx.workbook.application.calculate('Full');
ctx.executeAsync().then();
```

