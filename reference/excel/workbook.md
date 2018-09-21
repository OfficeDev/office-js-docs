# Workbook Object (JavaScript API for Excel)

Workbook is the top level object which contains related workbook objects such as worksheets, tables, ranges, etc.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|autoSave|bool|True if the workbook is in auto save mode. Read-only.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|calculationEngineVersion|int|Returns a number about the version of Excel Calculation Engine. Read-Only. Read-only.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|chartDataPointTrack|bool|True if all charts in the workbook are tracking the actual data points to which they are attached.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|isDirty|bool|True if no changes have been made to the specified workbook since it was last saved.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|name|string|Gets the workbook name. Read-only.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|previouslySaved|bool|True if the workbook has ever been saved locally or online. Read-only.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|readOnly|bool|True if the workbook is open in Read-only mode. Read-only.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|use1904DateSystem|bool|True if the workbook uses the 1904 date system.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|usePrecisionAsDisplayed|bool|True if calculations in this workbook will be done using only the precision of the numbers as they're displayed.|[beta](../requirement-sets/excel-api-requirement-sets.md)|

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|application|[Application](application.md)|Represents the Excel application instance that contains this workbook. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|bindings|[BindingCollection](bindingcollection.md)|Represents a collection of bindings that are part of the workbook. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|customXmlParts|[CustomXmlPartCollection](customxmlpartcollection.md)|Represents the collection of custom XML parts contained by this workbook. Read-only.|[1.5](../requirement-sets/excel-api-requirement-sets.md)|
|dataConnections|[DataConnectionCollection](dataconnectioncollection.md)|Represents all data connections in the workbook. Read-only.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|functions|[Functions](functions.md)|Represents a collection of worksheet functions that can be used for computation. Read-only.|[1.2](../requirement-sets/excel-api-requirement-sets.md)|
|names|[NamedItemCollection](nameditemcollection.md)|Represents a collection of workbook scoped named items (named ranges and constants). Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|pivotTables|[PivotTableCollection](pivottablecollection.md)|Represents a collection of PivotTables associated with the workbook. Read-only.|[1.3](../requirement-sets/excel-api-requirement-sets.md)|
|properties|[DocumentProperties](documentproperties.md)|Gets the workbook properties. Read-only.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|protection|[WorkbookProtection](workbookprotection.md)|Returns workbook protection object for a workbook. Read-only.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|settings|[SettingCollection](settingcollection.md)|Represents a collection of Settings associated with the workbook. Read-only.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|styles|[StyleCollection](stylecollection.md)|Represents a collection of styles associated with the workbook. Read-only.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|tables|[TableCollection](tablecollection.md)|Represents a collection of tables associated with the workbook. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|worksheets|[WorksheetCollection](worksheetcollection.md)|Represents a collection of worksheets associated with the workbook. Read-only.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[close(closeBehavior: CloseBehavior)](#closeclosebehavior-closebehavior)|void|Close current workbook.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[getActiveCell()](#getactivecell)|[Range](range.md)|Gets the currently active cell from the workbook.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|[getActiveChart()](#getactivechart)|[Chart](chart.md)|Gets the currently active chart in the workbook. If there is no active chart, will throw exception when invoke this statement|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[getActiveChartOrNullObject()](#getactivechartornullobject)|[Chart](chart.md)|Gets the currently active chart in the workbook. If there is no active chart, will return null object|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[getIsActiveCollabSession()](#getisactivecollabsession)|bool|True if the workbook is being edited by multiple users (co-authoring).|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[getRange(address: string)](#getrangeaddress-string)|[Range](range.md)|Gets the range object specified by the address or name.|[Design](../requirement-sets/excel-api-requirement-sets.md)|
|[getSelectedRange()](#getselectedrange)|[Range](range.md)|Gets the currently selected single range from the workbook. If there are multiple ranges selected, this method will throw an error.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[getSelectedRanges()](#getselectedranges)|[RangeAreas](rangeareas.md)|Gets the currently selected one or more ranges from the workbook. Unlike getSelectedRange(), this method returns a RangeAreas object that represents all the selected ranges.|[beta](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### close(closeBehavior: CloseBehavior)
Close current workbook.

#### Syntax
```js
workbookObject.close(closeBehavior);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|closeBehavior|CloseBehavior|Optional. workbook close behavior.|

#### Returns
void

### getActiveCell()
Gets the currently active cell from the workbook.

#### Syntax
```js
workbookObject.getActiveCell();
```

#### Parameters
None

#### Returns
[Range](range.md)

### getActiveChart()
Gets the currently active chart in the workbook. If there is no active chart, will throw exception when invoke this statement

#### Syntax
```js
workbookObject.getActiveChart();
```

#### Parameters
None

#### Returns
[Chart](chart.md)

### getActiveChartOrNullObject()
Gets the currently active chart in the workbook. If there is no active chart, will return null object

#### Syntax
```js
workbookObject.getActiveChartOrNullObject();
```

#### Parameters
None

#### Returns
[Chart](chart.md)

### getIsActiveCollabSession()
True if the workbook is being edited by multiple users (co-authoring).

#### Syntax
```js
workbookObject.getIsActiveCollabSession();
```

#### Parameters
None

#### Returns
bool

### getRange(address: string)
Gets the range object specified by the address or name.

#### Syntax
```js
workbookObject.getRange(address);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|address|string|The address or the name of the range. If the sheet name is not specified, the currently-active worksheet is assumed.|

#### Returns
[Range](range.md)

### getSelectedRange()
Gets the currently selected single range from the workbook. If there are multiple ranges selected, this method will throw an error.

#### Syntax
```js
workbookObject.getSelectedRange();
```

#### Parameters
None

#### Returns
[Range](range.md)

#### Examples

```js
Excel.run(function (ctx) { 
	var selectedRange = ctx.workbook.getSelectedRange();
	selectedRange.load('address');
	return ctx.sync().then(function() {
			console.log(selectedRange.address);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
### getSelectedRanges()
Gets the currently selected one or more ranges from the workbook. Unlike getSelectedRange(), this method returns a RangeAreas object that represents all the selected ranges.

#### Syntax
```js
workbookObject.getSelectedRanges();
```

#### Parameters
None

#### Returns
[RangeAreas](rangeareas.md)
