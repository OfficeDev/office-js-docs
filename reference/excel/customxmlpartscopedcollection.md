# CustomXmlPartScopedCollection Object (JavaScript API for Excel)

A scoped collection of custom XML parts.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|items|[CustomXmlPartScoped[]](customxmlpartscoped.md)|A collection of customXmlPartScoped objects. Read-only.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[getCount()](#getcount)|int|Gets the number of CustomXML parts in this collection.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(id: string)](#getitemid-string)|[CustomXmlPart](customxmlpart.md)|Gets a custom XML part based on its ID.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemOrNullObject(id: string)](#getitemornullobjectid-string)|[CustomXmlPart](customxmlpart.md)|Gets a custom XML part based on its ID.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getOnlyItem()](#getonlyitem)|[CustomXmlPart](customxmlpart.md)|If the collection contains exactly one item, this method returns it.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[getOnlyItemOrNullObject()](#getonlyitemornullobject)|[CustomXmlPart](customxmlpart.md)|If the collection contains exactly one item, this method returns it.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### getCount()
Gets the number of CustomXML parts in this collection.

#### Syntax
```js
customXmlPartScopedCollectionObject.getCount();
```

#### Parameters
None

#### Returns
int

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

### getItemOrNullObject(id: string)
Gets a custom XML part based on its ID.

#### Syntax
```js
customXmlPartScopedCollectionObject.getItemOrNullObject(id);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|id|string|ID of the object to be retrieved.|

#### Returns
[CustomXmlPart](customxmlpart.md)

### getOnlyItem()
If the collection contains exactly one item, this method returns it.

#### Syntax
```js
customXmlPartScopedCollectionObject.getOnlyItem();
```

#### Parameters
None

#### Returns
[CustomXmlPart](customxmlpart.md)

### getOnlyItemOrNullObject()
If the collection contains exactly one item, this method returns it.

#### Syntax
```js
customXmlPartScopedCollectionObject.getOnlyItemOrNullObject();
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
