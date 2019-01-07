# CommentReply Object (JavaScript API for Excel)

Represents a cell comment reply object in the workbook.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|content|string|GetSet the content.|[1.9](../requirement-sets/excel-api-requirement-sets.md)|
|id|string|Represents the comment reply identifier. Read-only.|[1.9](../requirement-sets/excel-api-requirement-sets.md)|
|isParent|bool|Represents whether it is a comment thread or reply. Always return false here. Read-only.|[1.9](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[delete()](#delete)|void|Deletes the comment reply.|[1.9](../requirement-sets/excel-api-requirement-sets.md)|
|[getParentComment()](#getparentcomment)|[Comment](comment.md)|Get its parent comment of this reply.|[1.9](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### delete()
Deletes the comment reply.

#### Syntax
```js
commentReplyObject.delete();
```

#### Parameters
None

#### Returns
void

### getParentComment()
Get its parent comment of this reply.

#### Syntax
```js
commentReplyObject.getParentComment();
```

#### Parameters
None

#### Returns
[Comment](comment.md)
