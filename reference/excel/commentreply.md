# CommentReply Object (JavaScript API for Excel)

Represents a cell comment reply object in the workbook.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|content|string|Gets or sets the content of the comment reply.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|id|string|Represents the comment reply identifier. Read-only.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|isParent|bool|Represents whether it is a comment thread or reply. Always return false here. Read-only.|[beta](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[delete()](#delete)|void|Deletes the comment reply.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[getParentComment()](#getparentcomment)|[Comment](comment.md)|Get its parent comment of this reply.|[beta](../requirement-sets/excel-api-requirement-sets.md)|

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

#### Examples

Delete a comment reply given reply ID

```js
Excel.run(async (context) => {
    var comment = context.workbook.comments.getItemByReplyID("{42D7DCA6-8FA5-4CA2-B089-107DFD534F00}");
    context.load(comment);
    return context.sync();

    comment.replies.getItem("{42D7DCA6-8FA5-4CA2-B089-107DFD534F00}").delete();
    return context.sync(); 
}).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

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
### Property access examples

Get a comment reply and change the content

```js
Excel.run(async (context) => {

    var comment = context.workbook.comments.getItemByReplyID("{42D7DCA6-8FA5-4CA2-B089-107DFD534F00}");
    context.load(comment);
    return context.sync();

    var reply = comment.replies.getItem("{42D7DCA6-8FA5-4CA2-B089-107DFD534F00}");
    reply.content = "new content for the reply";
    return context.sync();

}).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```
