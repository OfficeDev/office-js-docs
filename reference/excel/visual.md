# Visual Object (JavaScript API for Excel)

Represents a visual object in a workbook.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|id|string|The unique id of visual, not the guid of VisualDefinition. Read-only.|[ApiSet.InProgressFeatures.BiplatVisual](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|properties|[VisualPropertyCollection](visualpropertycollection.md)|Gets all properties of the visual. Read-only.|[ApiSet.InProgressFeatures.BiplatVisual](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[changeDataSource(dataSourceType: string, dataSourceContent: string)](#changedatasourcedatasourcetype-string-datasourcecontent-string)|void|Get the preview of a visual.|[ApiSet.InProgressFeatures.BiplatVisual](../requirement-sets/excel-api-requirement-sets.md)|
|[delete()](#delete)|void|Delete the visual.|[ApiSet.InProgressFeatures.BiplatVisual](../requirement-sets/excel-api-requirement-sets.md)|
|[getChildProperties(parentPropId: string)](#getchildpropertiesparentpropid-string)|[VisualPropertyCollection](visualpropertycollection.md)|Get the child properties of the specific parent property Id.|[ApiSet.InProgressFeatures.BiplatVisual](../requirement-sets/excel-api-requirement-sets.md)|
|[getDataControllerClient()](#getdatacontrollerclient)|[DataControllerClient](datacontrollerclient.md)|Get the DataControllerClient for the visual|[ApiSet.InProgressFeatures.BiplatVisual](../requirement-sets/excel-api-requirement-sets.md)|
|[getDataSource()](#getdatasource)|string|Gets a string represening the visual's current data source. e.g. Sheet1!$C$5:$D$7|[ApiSet.InProgressFeatures.BiplatVisual](../requirement-sets/excel-api-requirement-sets.md)|
|[getProperty(propName: string)](#getpropertypropname-string)|object|GetProperty|[ApiSet.InProgressFeatures.BiplatVisual](../requirement-sets/excel-api-requirement-sets.md)|
|[setProperty(propName: string, value: any)](#setpropertypropname-string-value-any)|void|SetProperty|[ApiSet.InProgressFeatures.BiplatVisual](../requirement-sets/excel-api-requirement-sets.md)|
|[setPropertyToDefault(propName: string)](#setpropertytodefaultpropname-string)|void|Returns true when the property's value is currently the default|[ApiSet.InProgressFeatures.BiplatVisual](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### changeDataSource(dataSourceType: string, dataSourceContent: string)
Get the preview of a visual.

#### Syntax
```js
visualObject.changeDataSource(dataSourceType, dataSourceContent);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|dataSourceType|string|The data source type of visual. e.g. xlFormula|
|dataSourceContent|string|The data source content. e.g. Sheet1!$C$5:$D$7|

#### Returns
void

### delete()
Delete the visual.

#### Syntax
```js
visualObject.delete();
```

#### Parameters
None

#### Returns
void

### getChildProperties(parentPropId: string)
Get the child properties of the specific parent property Id.

#### Syntax
```js
visualObject.getChildProperties(parentPropId);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|parentPropId|string|Optional. Parent property Id. Omitting this parameter will return the root-level properties.|

#### Returns
[VisualPropertyCollection](visualpropertycollection.md)

### getDataControllerClient()
Get the DataControllerClient for the visual

#### Syntax
```js
visualObject.getDataControllerClient();
```

#### Parameters
None

#### Returns
[DataControllerClient](datacontrollerclient.md)

### getDataSource()
Gets a string represening the visual's current data source. e.g. Sheet1!$C$5:$D$7

#### Syntax
```js
visualObject.getDataSource();
```

#### Parameters
None

#### Returns
string

### getProperty(propName: string)
GetProperty

#### Syntax
```js
visualObject.getProperty(propName);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|propName|string|...|

#### Returns
object

### setProperty(propName: string, value: any)
SetProperty

#### Syntax
```js
visualObject.setProperty(propName, value);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|propName|string|...|
|value|any|...|

#### Returns
void

### setPropertyToDefault(propName: string)
Returns true when the property's value is currently the default

#### Syntax
```js
visualObject.setPropertyToDefault(propName);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|propName|string|...|

#### Returns
void
