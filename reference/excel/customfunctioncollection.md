# CustomFunctionCollection Object (JavaScript API for Excel)

A collection of custom functions in Excel. You use this collection to register and unregister custom functions in Excel.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|items|[CustomFunction[]](customfunction.md)|A collection of customFunction objects. Read-only.|[ApiSet.CustomBase.CustomFunctions](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[deleteAll()](#deleteall)|void|Deletes all custom functions added by this add-in.|[ApiSet.CustomBase.CustomFunctions](../requirement-sets/excel-api-requirement-sets.md)|
|[getCount()](#getcount)|int|Gets the number of functions in the collection.|[ApiSet.CustomBase.CustomFunctions](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(name: string)](#getitemname-string)|[CustomFunction](customfunction.md)|Gets a function based on its name.|[ApiSet.CustomBase.CustomFunctions](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNullObject(name: string)](#getitemornullobjectname-string)|[CustomFunction](customfunction.md)|Gets a function based on its name.|[ApiSet.CustomBase.CustomFunctions](../requirement-sets/excel-api-requirement-sets.md)|
|[importFromWeb(metadataFormat: CustomFunctionMetadataFormat, metadataUrl: string, name: string)](#importfromwebmetadataformat-customfunctionmetadataformat-metadataurl-string-name-string)|[CustomFunction](customfunction.md)|Imports a new custom function to the workbook from web service metadata.|[ApiSet.CustomBase.CustomFunctions](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### deleteAll()
Deletes all custom functions added by this add-in.

#### Syntax
```js
customFunctionCollectionObject.deleteAll();
```

#### Parameters
None

#### Returns
void

### getCount()
Gets the number of functions in the collection.

#### Syntax
```js
customFunctionCollectionObject.getCount();
```

#### Parameters
None

#### Returns
int

### getItem(name: string)
Gets a function based on its name.

#### Syntax
```js
customFunctionCollectionObject.getItem(name);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|name|string|The name of the function to be retrieved.|

#### Returns
[CustomFunction](customfunction.md)

### getItemOrNullObject(name: string)
Gets a function based on its name.

#### Syntax
```js
customFunctionCollectionObject.getItemOrNullObject(name);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|name|string|The name of the function to be retrieved.|

#### Returns
[CustomFunction](customfunction.md)

### importFromWeb(metadataFormat: CustomFunctionMetadataFormat, metadataUrl: string, name: string)
Imports a new custom function to the workbook from web service metadata.

#### Syntax
```js
customFunctionCollectionObject.importFromWeb(metadataFormat, metadataUrl, name);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|metadataFormat|CustomFunctionMetadataFormat|The format of the metadata.|
|metadataUrl|string|The URL from where the metadata could be downloaded.|
|name|string|A new name to override the one from the metadata.|

#### Returns
[CustomFunction](customfunction.md)
