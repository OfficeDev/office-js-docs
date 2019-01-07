# CommentReplyCollection Object (JavaScript API for Excel)

Represents a collection of comment reply objects that are part of the comment.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|items|[CommentReply[]](commentreply.md)|A collection of commentReply objects. Read-only.|[beta](../requirement-sets/excel-api-requirement-sets.md)|

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[add(content: string, contentType: ContentType)](#addcontent-string-contenttype-contenttype)|[CommentReply](commentreply.md)|Creates a comment reply for comment.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[getCount()](#getcount)|int|Gets the number of comment replies in the collection.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(commentReplyId: string)](#getitemcommentreplyid-string)|[CommentReply](commentreply.md)|Returns a comment reply identified by its ID. Read-only.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[getItemAt(index: number)](#getitematindex-number)|[CommentReply](commentreply.md)|Gets a comment reply based on its position in the collection.|[beta](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### add(content: string, contentType: ContentType)
Creates a comment reply for comment.

#### Syntax
```js
commentReplyCollectionObject.add(content, contentType);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|content|string|The comment content.|
|contentType|ContentType|Optional. Optional. Type of the comment content|

#### Returns
[CommentReply](commentreply.md)

#### Examples
```js
Excel.run(async (context) => {
    var comment = context.workbook.comments.getItemByCell("sheet1!A1");
    context.load(comment);
    return context.sync();

    comment.replies.add("text of the reply");
    return context.sync();
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
### getCount()
Gets the number of comment replies in the collection.

#### Syntax
```js
commentReplyCollectionObject.getCount();
```

#### Parameters
None

#### Returns
int

### getItem(commentReplyId: string)
Returns a comment reply identified by its ID. Read-only.

#### Syntax
```js
commentReplyCollectionObject.getItem(commentReplyId);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|commentReplyId|string|The identifier for the comment reply.|

#### Returns
[CommentReply](commentreply.md)

### getItemAt(index: number)
Gets a comment reply based on its position in the collection.

#### Syntax
```js
commentReplyCollectionObject.getItemAt(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number|Index value of the object to be retrieved. Zero-indexed.|

#### Returns
[CommentReply](commentreply.md)
