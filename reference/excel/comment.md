# Comment Object (JavaScript API for Excel)

Represents a cell comment object in the workbook.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|content|string|Gets or sets the content of the comment.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|id|string|Represents the comment identifier. Read-only.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|isParent|bool|Represents whether it is a comment thread or reply. Always return true here. Read-only.|[beta](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|replies|[CommentReplyCollection](commentreplycollection.md)|Represents a collection of reply objects associated with the comment. Read-only.|[beta](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[delete()](#delete)|void|Deletes the comment thread.|[beta](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### delete()
Deletes the comment thread.

#### Syntax
```js
commentObject.delete();
```

#### Parameters
None

#### Returns
void

#### Examples

```js
Excel.run(async (context) => {
	context.workbook.comments.getItem("{42D7DCA6-8FA5-4CA2-B089-107DFD534F00}").delete();
	return context.sync();
}).catch(function(error) {
	console.log("Error: " + error);
	if (error instanceof OfficeExtension.Error) {
		console.log("Debug info: " + JSON.stringify(error.debugInfo));
	}
});
```
### Property access examples
```js
Excel.run(async (context) => {
    var comment = context.workbook.comments.getItemAt(0);
    comment.content = "new content for the comment";
    return context.sync();
}).catch(function(error) {
	console.log("Error: " + error);
	if (error instanceof OfficeExtension.Error) {
		console.log("Debug info: " + JSON.stringify(error.debugInfo));
	}
});
```

