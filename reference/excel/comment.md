# Comment Object (JavaScript API for Excel)

Represents a cell comment object in the workbook.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|content|string|GetSet the content.|[1.9](../requirement-sets/excel-api-requirement-sets.md)|
|id|string|Represents the comment identifier. Read-only.|[1.9](../requirement-sets/excel-api-requirement-sets.md)|
|isParent|bool|Represents whether it is a comment thread or reply. Always return true here. Read-only.|[1.9](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|replies|[CommentReplyCollection](commentreplycollection.md)|Represents a collection of reply objects associated with the comment. Read-only.|[1.9](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[delete()](#delete)|void|Deletes the comment thread.|[1.9](../requirement-sets/excel-api-requirement-sets.md)|

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
