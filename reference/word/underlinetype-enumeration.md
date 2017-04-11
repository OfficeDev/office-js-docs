# UnderlineType Enumeration (JavaScript API for Word)

Represents a font's underline type.

_Applies to: Word 2016, Word for iPad, Word for Mac_

 ## Properties
| Enumeration                   | Value	            |Description    |
|:------------------------------|:------------------|:--------------|
| UnderlineType.none            | "None"            | No underline |
| UnderlineType.single          | "Single"          | Single underline |
| UnderlineType.word            | "Word"            | Underline each word (excludes spaces) |
| UnderlineType.double          | "Double"          | Double underlined |
| UnderlineType.dotted          | "Dotted"          | Dotted underline |
| UnderlineType.hidden          | "Hidden"          | Hidden underline |
| UnderlineType.thick           | "Thick"           | Thicker version of Single underline |
| UnderlineType.dashLine        | "DashLine"        | Dashed underline |
| UnderlineType.dotLine         | "DotLine"         | Tightly kerned dotted underline |
| UnderlineType.dotDashLine     | "DotDashLine"     | Dot-Dash-Dot underline|
| UnderlineType.twoDotDashLine  | "TwoDotDashLine"  | Dot-Dot-Dash-Dot-Dot underline |
| UnderlineType.wave            | "Wave"            | Wavy underline (aka "squiggle" underline ) |

## Example

Underline format text
```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a range proxy object for the current selection.
    var selection = context.document.getSelection();

    // Queue a commmand to underline the current selection.
    selection.font.underline = Word.UnderlineType.single;

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log('The selection now has an underline style.');
    });
})
.catch(function (error) {
    console.log('Error: ' + JSON.stringify(error));
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```