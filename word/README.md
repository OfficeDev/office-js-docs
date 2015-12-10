# Word JavaScript APIs

Welcome to the Word JavaScript API documentation repository. Here you'll find what you'll need to create the next generation of Word add-ins in Office 2016 for Windows (if you can't find it, open an issue). The new Word JavaScript APIs provide Word-specific functionality related to documents, paragraphs, content controls, and other common Word objects. This API complements the functionality of our existing Office.js library. 

This documentation is [published on MSDN](https://msdn.microsoft.com/EN-US/library/office/mt616496.aspx). 

## What's new in the WordAPI 1.3 requirement set
This branch contains the new APIs that our team is working on. We are plan to ship these changes in the next few months. This is a great time to give feedback on these APIs.

### 1. Granular access to [range](/word/word-add-ins-javascript-reference/range.md) objects. 

We are adding the following members to the [body](/word/word-add-ins-javascript-reference/body.md), [contentcontrol](/word/word-add-ins-javascript-reference/contentcontrol.md), [paragraph](/word/word-add-ins-javascript-reference/paragraph.md) and [range](/word/word-add-ins-javascript-reference/range.md) objects:

```js
    
    // Get child range.
    
	Range Body.GetChildRange(RangeOrigin rangeOrigin, int length);
	Range ContentControl.GetChildRange(RangeOrigin rangeOrigin, int length);
	Range Paragraph.GetChildRange(RangeOrigin rangeOrigin, int length);
	Range Range.GetChildRange(RangeOrigin rangeOrigin, int length);

    // Granular access to ranges via delimiters.

	RangeCollection Body.GetRanges(string[] delimiters, [Optional] bool excludeDelimiter, [Optional] bool trimSpace, [Optional] bool excludeHiddenChars);
	RangeCollection ContentControl.GetRanges(string[] delimiters, [Optional] bool excludeDelimiter, [Optional] bool trimSpace, [Optional] bool excludeEop, [Optional] bool within);
	RangeCollection Paragraph.GetRanges(string[] delimiters, [Optional] bool excludeDelimiter, [Optional] bool trimSpace);
	RangeCollection Range.GetRanges(string[] delimiters, [Optional] bool excludeDelimiter, [Optional] bool trimSpace, [Optional] bool excludeHiddenChars, [Optional] bool within);

    // New range object features: range inclusion comparison, expansion, and range boundary adjustments.

	bool Range.HasRange(Range range, [Optional] bool isSubset); // Determine whether a range includes another range.
	void Range.ExpandTo(Range range); // Expand a range to include the bounds of another range.
	void Range.Adjust(int startAdjust, int endAdjust); // Adjust the start and end of the range by the specified number of characters. 

```


### 2. Create and open a new Word Document. 

This feature enables you to create, change, open, and make changes to a Word document (.docx) before it is displayed in the Word UI.

### 3. Strongly typed List objects.

### 4. Strongly typed Table objects.

## Try it out (only available for released features)

We've been working on a Snippet Explorer to let you browse through common code snippets and learn how the new APIs work. Give it a try. The code snippets referenced by the Snippet Explorer are available [here](https://officesnippetexplorer.azurewebsites.net/#/snippets/word). 

## Give feedback

We need it, you want to give it. Feedback is much easier to give now that we're on GitHub. Check out the docs and let us know about any [issues](https://github.com/OfficeDev/office-js-docs/issues) you find in those docs by submitting issues directly in this repository. Let us know what you think about the APIs and the general programming experience. 

We suggest you use the tags [office-js] and [word] on StackOverflow for asking questions to the community.
