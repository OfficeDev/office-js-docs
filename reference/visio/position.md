# Position object (JavaScript API for Visio)

Applies to: _Visio Online_

Represents the Position of the object in the view.

## Properties

| Property	   | Type	|Description|
|:---------------|:--------|:----------|
|x|int|An integer that specifies the x-coordinate of the object, which is the signed value of the distance in pixels from the viewport's center to the left boundary of the page.|
|y|int|An integer that specifies the y-coordinate of the object, which is the signed value of the distance in pixels from the viewport's center to the top boundary of the page.|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method Details


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
