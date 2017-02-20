# ShapeMouseEnterEventArgs Object (JavaScript API for OneNote)

_Applies to: OneNote Online_
_Note: This API is in preview_

Provides information about the shape that raised the MouseEnter event.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|pageName|string|Gets the name of the page which has the shape object that raised the MouseEnter event.|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|string|[public](public.md)|Gets the shape object that raised the MouseEnter event.|

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
|:---------------|:--------|:----------|
|param|object|Optional. Accepts parameter and relationship names as delimited string or an array. Or, provide [loadOption](loadoption.md) object.|

#### Returns
void
