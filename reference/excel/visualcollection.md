# VisualCollection Object (JavaScript API for Excel)

A collection of all the visuals on a worksheet.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|items|[Visual[]](visual.md)|A collection of visual objects. Read-only.|[1.9](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[add(visualDefinitionGuid : string, dataSourceType: string, dataSourceContent: string)](#addvisualdefinitionguid--string-datasourcetype-string-datasourcecontent-string)|[Visual](visual.md)|Creates a new visual.|[1.9](../requirement-sets/excel-api-requirement-sets.md)|
|[getDefinitions()](#getdefinitions)|[VisualDefinition[]](visualdefinition[].md)|Gets all visual definitions.|[1.9](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(id: string)](#getitemid-string)|[Visual](visual.md)|Get item.|[1.9](../requirement-sets/excel-api-requirement-sets.md)|
|[getPreview(visualDefinitionGuid : string, width : number, height : number, dpi : number)](#getpreviewvisualdefinitionguid--string-width--number-height--number-dpi--number)|[System.IO.Stream](system.io.stream.md)|Get the preview of a visual.|[1.9](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### add(visualDefinitionGuid : string, dataSourceType: string, dataSourceContent: string)
Creates a new visual.

#### Syntax
```js
visualCollectionObject.add(visualDefinitionGuid , dataSourceType, dataSourceContent);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|visualDefinitionGuid |string|The guid of a VisualDefinition, not the id for an instance of a Visual.|
|dataSourceType|string|Optional. The data source type of visual. e.g. xlFormula|
|dataSourceContent|string|Optional. The data source content|

#### Returns
[Visual](visual.md)

### getDefinitions()
Gets all visual definitions.

#### Syntax
```js
visualCollectionObject.getDefinitions();
```

#### Parameters
None

#### Returns
[VisualDefinition[]](visualdefinition[].md)

### getItem(id: string)
Get item.

#### Syntax
```js
visualCollectionObject.getItem(id);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|id|string|ID.|

#### Returns
[Visual](visual.md)

### getPreview(visualDefinitionGuid : string, width : number, height : number, dpi : number)
Get the preview of a visual.

#### Syntax
```js
visualCollectionObject.getPreview(visualDefinitionGuid , width , height , dpi );
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|visualDefinitionGuid |string|The guid of a VisualDefinition, not the id for an instance of a Visual.|
|width |number|The width of the preview.|
|height |number|The height of the preview.|
|dpi |number|The dpi setting.|

#### Returns
[System.IO.Stream](system.io.stream.md)
