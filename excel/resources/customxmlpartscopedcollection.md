# CustomXmlPartScopedCollection Object (JavaScript API for Excel)

_Excel 2016, Excel Online, Excel for iPad, Excel for Mac_

A scoped collection of custom XML parts.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|items|[CustomXmlPartScoped[]](customxmlpartscoped.md)|A collection of customXmlPartScoped objects. Read-only.|1.3||

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[getItem(id: string)](#getitemid-string)|[CustomXmlPart](customxmlpart.md)|Gets a custom XML part based on its ID.|1.3|
|[getOnlyItemOrNull()](#getonlyitemornull)|[CustomXmlPart](customxmlpart.md)|If the collection contains exactly one item, this method returns it.|1.3|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|1.1|

## Method Details


### getItem(id: string)
Gets a custom XML part based on its ID.

#### Syntax
```js
customXmlPartScopedCollectionObject.getItem(id);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|id|string|ID of the object to be retrieved.|

#### Returns
[CustomXmlPart](customxmlpart.md)

### getOnlyItemOrNull()
If the collection contains exactly one item, this method returns it.

#### Syntax
```js
customXmlPartScopedCollectionObject.getOnlyItemOrNull();
```

#### Parameters
None

#### Returns
[CustomXmlPart](customxmlpart.md)

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
