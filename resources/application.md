# Application

Represents the Excel application that manages the workbook.

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
|[calculate()](#calculate)| Void |Perform calculation on the workbook or application.| |

### Get Application

Get the properties of the Application object.

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
`calculationType` | String | Optional. Specifies the calculation type to use. Possible values are: `ReCalculate`: Performs normal calculation by calculating all the formulas in the workbook, `Full`: Forces a full calculation of the data, `FullRebuild`: Forces a full calculation of the data and rebuilds the dependencies. This option is similar to re-entering all formulas. Note: If calculationType is not specified, the 'ReCalculate' option is used by default.

#### Returns

Nothing

#### Examples 

```js
var ctx = new Excel.ExcelClientContext();
ctx.workbook.application.calculate('Full');
ctx.executeAsync().then();
```

