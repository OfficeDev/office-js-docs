# PageCollection Object (JavaScript API for Visio)

_Visio Online_

Represents a collection of Page objects that are part of the document.

## Properties

| Property	   | Type	|Description| Req. Set| Feedback|
|:---------------|:--------|:----------|:----|:---|
|items|[Page[]](page.md)|A collection of page objects. Read-only.|1.1|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-pageCollection-items)|

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set| Feedback|
|:---------------|:--------|:----------|:----|:---|
|[getCount()](#getcount)|int|Gets the number of pages in the collection.|1.1|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-pageCollection-getCount)|
|[getItem(key: number or string)](#getitemkey-number-or-string)|[Page](page.md)|Gets a page using its key (name or Id).|1.1|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-pageCollection-getItem)|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|1.1|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=Visio-pageCollection-load)|

## Method Details


### getCount()
Gets the number of pages in the collection.

#### Syntax
```js
pageCollectionObject.getCount();
```

#### Parameters
None

#### Returns
int

### getItem(key: number or string)
Gets a page using its key (name or Id).

#### Syntax
```js
pageCollectionObject.getItem(key);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|key|number or string|Key is the name or Id of the page to be retrieved.|

#### Returns
[Page](page.md)

#### Examples
```js
Visio.run(function (ctx) { 
	var pageName = 'Page-1';
	var page = ctx.document.pages.getItem(pageName);
	page.activate();
	return ctx.sync();
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

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
