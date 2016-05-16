# WorksheetCollection Object (JavaScript API for Excel)

_Excel 2016, Excel Online, Excel for iPad, Excel for Mac_

Represents a collection of worksheet objects that are part of the workbook.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|items|[Worksheet[]](worksheet.md)|A collection of worksheet objects. Read-only.|1.1||

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[add(name: string)](#addname-string)|[Worksheet](worksheet.md)|Adds a new worksheet to the workbook. The worksheet will be added at the end of existing worksheets. If you wish to activate the newly added worksheet, call ".activate() on it.|1.1|
|[getActiveWorksheet()](#getactiveworksheet)|[Worksheet](worksheet.md)|Gets the currently active worksheet in the workbook.|1.1|
|[getItem(key: string)](#getitemkey-string)|[Worksheet](worksheet.md)|Gets a worksheet object using its Name or ID.|1.1|
|[getItemOrNull(key: string)](#getitemornullkey-string)|[Worksheet](worksheet.md)|Gets a worksheet object using its Name or ID. If the worksheet does not exist, the returned object's isNull property will be true.|1.3|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|1.1|

## Method Details


### add(name: string)
Adds a new worksheet to the workbook. The worksheet will be added at the end of existing worksheets. If you wish to activate the newly added worksheet, call ".activate() on it.

#### Syntax
```js
worksheetCollectionObject.add(name);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|name|string|Optional. The name of the worksheet to be added. If specified, name should be unqiue. If not specified, Excel determines the name of the new worksheet.|

#### Returns
[Worksheet](worksheet.md)

### getActiveWorksheet()
Gets the currently active worksheet in the workbook.

#### Syntax
```js
worksheetCollectionObject.getActiveWorksheet();
```

#### Parameters
None

#### Returns
[Worksheet](worksheet.md)

### getItem(key: string)
Gets a worksheet object using its Name or ID.

#### Syntax
```js
worksheetCollectionObject.getItem(key);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|key|string|The Name or ID of the worksheet.|

#### Returns
[Worksheet](worksheet.md)

### getItemOrNull(key: string)
Gets a worksheet object using its Name or ID. If the worksheet does not exist, the returned object's isNull property will be true.

#### Syntax
```js
worksheetCollectionObject.getItemOrNull(key);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|key|string|The Name or ID of the worksheet.|

#### Returns
[Worksheet](worksheet.md)

### load(param: object)
Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.

#### Syntax
```js
object.load(param);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|param|object|Optional. Accepts parameter and relationship names as delimited string or an array. Or, provide [loadOption](loadoption.md) object.|

#### Returns
void
