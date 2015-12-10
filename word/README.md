# Word JavaScript APIs

Welcome to the Word JavaScript API documentation repository. Here you'll find what you'll need to create the next generation of Word add-ins in Office 2016 for Windows (if you can't find it, open an issue). The new Word JavaScript APIs provide Word-specific functionality related to documents, paragraphs, content controls, and other common Word objects. This API complements the functionality of our existing Office.js library. 

This documentation is [published on MSDN](https://msdn.microsoft.com/EN-US/library/office/mt616496.aspx). 

## What's new
As indicated, this branch is holding the  API additions our team is working on right now and we are planning to ship in the next few months.

###1. Granular access to ranges. 

We are adding the following members to the [body](/word-add-ins-javascript-reference/body.md), contentControl, paragraph and range objects:


```js
//getting a child range.
	Range Body.GetChildRange(RangeOrigin rangeOrigin, int length);
	Range ContentControl.GetChildRange(RangeOrigin rangeOrigin, int length);
	Range Paragraph.GetChildRange(RangeOrigin rangeOrigin, int length);
	Range Range.GetChildRange(RangeOrigin rangeOrigin, int length);

//granular access to ranges via delimiters...

	RangeCollection Body.GetRanges(string[] delimiters, [Optional] bool excludeDelimiter, [Optional] bool trimSpace, [Optional] bool excludeHiddenChars);
	RangeCollection ContentControl.GetRanges(string[] delimiters, [Optional] bool excludeDelimiter, [Optional] bool trimSpace, [Optional] bool excludeEop, [Optional] bool within);
	RangeCollection Paragraph.GetRanges(string[] delimiters, [Optional] bool excludeDelimiter, [Optional] bool trimSpace);
	RangeCollection Range.GetRanges(string[] delimiters, [Optional] bool excludeDelimiter, [Optional] bool trimSpace, [Optional] bool excludeHiddenChars, [Optional] bool within);

//other range capabilities, inclusion, expansion, adjustment...

	bool Range.HasRange(Range range, [Optional] bool isSubset);
	void Range.ExpandTo(Range range);
	void Range.Adjust(int startAdjust, int endAdjust);


```


###2. The capability of creating and openening a separate Word Document.
###3. The ability to get and create Lists 
###4. Strongly type objects to allow Table manipulation.

## Try it out

We've been working on a Snippet Explorer to let you browse through common code snippets and learn how the new APIs work. Give it a try. The code snippets referenced by the Snippet Explorer are available [here](https://officesnippetexplorer.azurewebsites.net/#/snippets/word). 

## Give feedback

We need it, you want to give it. Feedback is much easier to give now that we're on GitHub. Check out the docs and let us know about any [issues](https://github.com/OfficeDev/office-js-docs/issues) you find in those docs by submitting issues directly in this repository. Let us know what you think about the APIs and the general programming experience. 

We suggest you use the tags [office-js] and [word] on StackOverflow for asking questions to the community.
