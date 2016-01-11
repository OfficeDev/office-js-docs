# Word JavaScript APIs

Welcome to the Word JavaScript API documentation repository. Here you'll find what you'll need to create the next generation of Word add-ins in Office 2016 for Windows (if you can't find it, open an issue). The new Word JavaScript APIs provide Word-specific functionality related to documents, paragraphs, content controls, and other common Word objects. This API complements the functionality of our existing Office.js library. 

This documentation is [published on MSDN](https://msdn.microsoft.com/EN-US/library/office/mt616496.aspx). 

## Introduction to Word JS APIs 1.3 
This branch contains the new APIs that our team is working on. We are plan to ship these changes in the next few months. This is a great time to give feedback on these APIs.

This section describes the new set of Word JavaScript APIs that are being planned for the next release (Requirement Set 1.3). Please review and provide your feedback. Provide your feedback by opening new issues in GitHub using the links in the table. 

_**Note**: The listed features are still under the design and review phase and are not yet available as part of the product. The final design is subject to change. Once the feature is made available, the final specification will be published as part of the master repository._


| New feature	   | Description	| Give feedback|
|:---------------|:--------|:----------|
|Get child ranges| 	Range Body.GetChildRange(RangeOrigin rangeOrigin, int length); <br/><br/>Range ContentControl.GetChildRange(RangeOrigin rangeOrigin, int length); <br/><br/>	Range Paragraph.GetChildRange(RangeOrigin rangeOrigin, int length); <br/><br/> Range Range.GetChildRange(RangeOrigin rangeOrigin, int length); | _[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=WordJs-1.3-OpenSpec-GetChildRanges)_|
|Get ranges via deliliters | RangeCollection Body.GetRanges(string\[] delimiters, \[Optional] bool excludeDelimiters, \[Optional] bool trimWhite, \[Optional] bool excludeEndingMarks); <br/><br/> 	RangeCollection ContentControl.GetRanges(string\[] delimiters, \[Optional] bool excludeDelimiters, \[Optional] bool trimWhite, \[Optional] bool excludeEndingMarks, \[Optional] bool within); <br/><br/> 	RangeCollection Paragraph.GetRanges(string\[] delimiters, \[Optional] bool excludeDelimiters, \[Optional] bool trimWhite); <br/><br/> 	RangeCollection Range.GetRanges(string\[] delimiters, \[Optional] bool excludeDelimiters, \[Optional] bool trimWhite, \[Optional] bool excludeEndingMarks, \[Optional] bool within); | _[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=WordJs-1.3-OpenSpec-GetRangesViaDelimiters)_|
|Range comparison and expansion| 	bool Range.HasRange(Range range, \[Optional] bool isSubset);  <br/><br/> void Range.ExpandTo(Range range); <br/><br/> void Range.Adjust(int startAdjust, int endAdjust);  | _[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=WordJs-1.3-OpenSpec-RangeCompareAndExpansion)_|
|Create and open a new Word Document| 	This feature enables you to create, change, open, and make changes to a Word document (.docx) before it is displayed in the Word UI. | _[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=WordJs-1.3-OpenSpec-NewWordDocument)_|
|Strongly typed List objects| 	This gives you access to ordered and unordered list objects. | _[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=WordJs-1.3-OpenSpec-NewListObject)_|
|Strongly typed Table objects| 	This gives you access to ordered and unordered list objects. | _[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=WordJs-1.3-OpenSpec-NewTableObject)_|



## Try it out (only available for released features)

We've been working on a Snippet Explorer to let you browse through common code snippets and learn how the new APIs work. Give it a try. The code snippets referenced by the Snippet Explorer are available [here](https://officesnippetexplorer.azurewebsites.net/#/snippets/word). 

## Give feedback

We need it, you want to give it. Feedback is much easier to give now that we're on GitHub. Check out the docs and let us know about any [issues](https://github.com/OfficeDev/office-js-docs/issues) you find in those docs by submitting issues directly in this repository. Let us know what you think about the APIs and the general programming experience. 

We suggest you use the tags [office-js] and [word] on StackOverflow for asking questions to the community.
