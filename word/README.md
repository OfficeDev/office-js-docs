# Word JavaScript APIs

Welcome to the Word JavaScript API documentation repository. Here you'll find what you'll need to create the next generation of Word add-ins in Office 2016 for Windows (if you can't find it, open an issue). The new Word JavaScript APIs provide Word-specific functionality related to documents, paragraphs, content controls, and other common Word objects. This API complements the functionality of our existing Office.js library. 

This documentation is [published on MSDN](https://msdn.microsoft.com/EN-US/library/office/mt616496.aspx). 

## Introduction to Word JS APIs 1.3 
This branch contains the new APIs that our team is working on. We are plan to ship these changes in the next few months. This is a great time to give feedback on these APIs.

This section describes the new set of Word JavaScript APIs that are being planned for the next release (Requirement Set 1.3). Please review and provide your feedback. Provide your feedback by opening new issues in GitHub using the links in the table. 

_**Note**: The listed features are still under the design and review phase and are not yet available as part of the product. The final design is subject to change. Once the feature is made available, the final specification will be published as part of the master repository._

### FEATURES

**Resource name:** [resources/application](application.md) </br>
**What's new:** Method **[createDocument(base64File: string)](resources/application.md#createdocumentbase64file-string)** returning **[Document](resources/document.md)** </br>
**Description:** Creates a new document by using a base64 encoded .docx file. </br>
**Available in requirement set:** WordApiDesktop, 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=application-createDocument)_ </br>

**Resource name:** [body](resources/body.md) </br>
**What's new:** Relationship **lists** of type **[ListCollection](listcollection.md)** </br>
**Description:** Gets the collection of list objects in the body. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=body-lists)_ </br>

**Resource name:** [body](resources/body.md) </br>
**What's new:** Relationship **parentBody** of type **[Body](body.md)** </br>
**Description:** Gets the parent body of the body. For example, a table cell body's parent body could be a header. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=body-parentBody)_ </br>

**Resource name:** [body](resources/body.md) </br>
**What's new:** Relationship **tables** of type **[TableCollection](tablecollection.md)** </br>
**Description:** Gets the collection of table objects in the body. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=body-tables)_ </br>

**Resource name:** [body](resources/body.md) </br>
**What's new:** Relationship **type** of type **[BodyType](bodytype.md)** </br>
**Description:** Gets the type of the body. The type can be 'MainDoc', 'Section', 'Header', 'Footer', or 'TableCell'. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=body-type)_ </br>

**Resource name:** [resources/body](body.md) </br>
**What's new:** Method **[getRange(rangeLocation: RangeLocation)](resources/body.md#getrangerangelocation-rangelocation)** returning **[Range](resources/range.md)** </br>
**Description:** Gets the whole body, or the starting or ending point of the body, as a range. </br>
**Available in requirement set:** 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=body-getRange)_ </br>

**Resource name:** [resources/body](body.md) </br>
**What's new:** Method **[insertTable(rowCount: number, columnCount: number, insertLocation: InsertLocation, values: string[][])](resources/body.md#inserttablerowcount-number-columncount-number-insertlocation-insertlocation-values-string)** returning **[Table](resources/table.md)** </br>
**Description:** Inserts a table with the specified number of rows and columns. The insertLocation value can be 'Start' or 'End'. </br>
**Available in requirement set:** 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=body-insertTable)_ </br>

**Resource name:** [contentControl](resources/contentcontrol.md) </br>
**What's new:** Relationship **lists** of type **[ListCollection](listcollection.md)** </br>
**Description:** Gets the collection of list objects in the content control. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=contentControl-lists)_ </br>

**Resource name:** [contentControl](resources/contentcontrol.md) </br>
**What's new:** Relationship **parentTable** of type **[Table](table.md)** </br>
**Description:** Gets the table that contains the content control. Returns null if it is not contained in a table. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=contentControl-parentTable)_ </br>

**Resource name:** [contentControl](resources/contentcontrol.md) </br>
**What's new:** Relationship **parentTableCell** of type **[TableCell](tablecell.md)** </br>
**Description:** Gets the table cell that contains the content control. Returns null if it is not contained in a table cell. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=contentControl-parentTableCell)_ </br>

**Resource name:** [contentControl](resources/contentcontrol.md) </br>
**What's new:** Relationship **subtype** of type **[ContentControlType](contentcontroltype.md)** </br>
**Description:** Gets the content control subtype. The subtype can be 'RichTextInline', 'RichTextParagraphs', 'RichTextTableCell', 'RichTextTableRow' and 'RichTextTable' for rich text content controls. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=contentControl-subtype)_ </br>

**Resource name:** [contentControl](resources/contentcontrol.md) </br>
**What's new:** Relationship **tables** of type **[TableCollection](tablecollection.md)** </br>
**Description:** Gets the collection of table objects in the content control. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=contentControl-tables)_ </br>

**Resource name:** [resources/contentControl](contentcontrol.md) </br>
**What's new:** Method **[getRange(rangeLocation: RangeLocation)](resources/contentcontrol.md#getrangerangelocation-rangelocation)** returning **[Range](resources/range.md)** </br>
**Description:** Gets the whole content control, or the starting or ending point of the content control, as a range. </br>
**Available in requirement set:** 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=contentControl-getRange)_ </br>

**Resource name:** [resources/contentControl](contentcontrol.md) </br>
**What's new:** Method **[getTextRanges(punctuationMarks: string[], trimSpacing: bool)](resources/contentcontrol.md#gettextrangespunctuationmarks-string-trimspacing-bool)** returning **[RangeCollection](resources/rangecollection.md)** </br>
**Description:** Gets the text ranges in the content control by using punctuation marks andor space character. </br>
**Available in requirement set:** 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=contentControl-getTextRanges)_ </br>

**Resource name:** [resources/contentControl](contentcontrol.md) </br>
**What's new:** Method **[insertTable(rowCount: number, columnCount: number, insertLocation: InsertLocation, values: string[][])](resources/contentcontrol.md#inserttablerowcount-number-columncount-number-insertlocation-insertlocation-values-string)** returning **[Table](resources/table.md)** </br>
**Description:** Inserts a table with the specified number of rows and columns into, or next to, a content control. The insertLocation value can be 'Start', 'End', 'Before' or 'After'. </br>
**Available in requirement set:** 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=contentControl-insertTable)_ </br>

**Resource name:** [resources/contentControl](contentcontrol.md) </br>
**What's new:** Method **[split(delimiters: string[], multiParagraphs: bool, trimDelimiters: bool, trimSpacing: bool)](resources/contentcontrol.md#splitdelimiters-string-multiparagraphs-bool-trimdelimiters-bool-trimspacing-bool)** returning **[RangeCollection](resources/rangecollection.md)** </br>
**Description:** Splits the content control into child ranges by using delimiters. </br>
**Available in requirement set:** 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=contentControl-split)_ </br>

**Resource name:** [resources/contentControlCollection](contentcontrolcollection.md) </br>
**What's new:** Method **[getByTypes(types: ContentControlType[])](resources/contentcontrolcollection.md#getbytypestypes-contentcontroltype)** returning **[ContentControlCollection](resources/contentcontrolcollection.md)** </br>
**Description:** Gets the content controls that have the specified types andor subtypes. </br>
**Available in requirement set:** WordApiDesktop, 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=contentControlCollection-getByTypes)_ </br>

**Resource name:** [resources/document](document.md) </br>
**What's new:** Method **[open()](resources/document.md#open)** returning **void** </br>
**Description:** Open the document. </br>
**Available in requirement set:** WordApiDesktop, 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=document-open)_ </br>

**Resource name:** [font](resources/font.md) </br>
**What's new:** Property **doubleStrikeThrough** of type **bool** </br>
**Description:** Gets or sets a value that indicates whether the font has a double strike through. True if the font is formatted as double strikethrough text, otherwise, false. </br>
**Available in requirement set:** WordApiDesktop, 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=font-doubleStrikeThrough)_ </br>

**Resource name:** [inlinePicture](resources/inlinepicture.md) </br>
**What's new:** Relationship **imageFormat** of type **[ImageFormat](imageformat.md)** </br>
**Description:** Gets the format of the inline image. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=inlinePicture-imageFormat)_ </br>

**Resource name:** [inlinePicture](resources/inlinepicture.md) </br>
**What's new:** Relationship **next** of type **[InlinePicture](inlinepicture.md)** </br>
**Description:** Gets the next inline image. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=inlinePicture-next)_ </br>

**Resource name:** [inlinePicture](resources/inlinepicture.md) </br>
**What's new:** Relationship **parentTable** of type **[Table](table.md)** </br>
**Description:** Gets the table that contains the inline image. Returns null if it is not contained in a table. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=inlinePicture-parentTable)_ </br>

**Resource name:** [inlinePicture](resources/inlinepicture.md) </br>
**What's new:** Relationship **parentTableCell** of type **[TableCell](tablecell.md)** </br>
**Description:** Gets the table cell that contains the inline image. Returns null if it is not contained in a table cell. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=inlinePicture-parentTableCell)_ </br>

**Resource name:** [resources/inlinePicture](inlinepicture.md) </br>
**What's new:** Method **[getRange(rangeLocation: RangeLocation)](resources/inlinepicture.md#getrangerangelocation-rangelocation)** returning **[Range](resources/range.md)** </br>
**Description:** Gets the picture, or the starting or ending point of the picture, as a range. </br>
**Available in requirement set:** 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=inlinePicture-getRange)_ </br>

**Resource name:** [inlinePictureCollection](resources/inlinepicturecollection.md) </br>
**What's new:** Relationship **first** of type **[InlinePicture](inlinepicture.md)** </br>
**Description:** Gets the first inline image in this collection. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=inlinePictureCollection-first)_ </br>

**Resource name:** [list](resources/list.md) </br>
**What's new:** Property **id** of type **int** </br>
**Description:** Gets the list's id. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=list-id)_ </br>

**Resource name:** [list](resources/list.md) </br>
**What's new:** Relationship **paragraphs** of type **[ParagraphCollection](paragraphcollection.md)** </br>
**Description:** A collection containing the paragraphs in this list. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=list-paragraphs)_ </br>

**Resource name:** [resources/list](list.md) </br>
**What's new:** Method **[insertParagraph(paragraphText: string, insertLocation: InsertLocation)](resources/list.md#insertparagraphparagraphtext-string-insertlocation-insertlocation)** returning **[Paragraph](resources/paragraph.md)** </br>
**Description:** Inserts a paragraph at the specified location. The insertLocation value can be 'Start', 'End', 'Before' or 'After'. </br>
**Available in requirement set:** 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=list-insertParagraph)_ </br>

**Resource name:** [listCollection](resources/listcollection.md) </br>
**What's new:** Property **items** of type **[List[]](list.md)** </br>
**Description:** A collection of list objects. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=listCollection-items)_ </br>

**Resource name:** [listCollection](resources/listcollection.md) </br>
**What's new:** Relationship **first** of type **[List](list.md)** </br>
**Description:** Gets the first list in this collection. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=listCollection-first)_ </br>

**Resource name:** [resources/listCollection](listcollection.md) </br>
**What's new:** Method **[getById(id: number)](resources/listcollection.md#getbyidid-number)** returning **[List](resources/list.md)** </br>
**Description:** Gets a list by its identifier. </br>
**Available in requirement set:** 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=listCollection-getById)_ </br>

**Resource name:** [resources/listCollection](listcollection.md) </br>
**What's new:** Method **[getItem(index: number)](resources/listcollection.md#getitemindex-number)** returning **[List](resources/list.md)** </br>
**Description:** Gets a list object by its index in the collection. </br>
**Available in requirement set:** 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=listCollection-getItem)_ </br>

**Resource name:** [paragraph](resources/paragraph.md) </br>
**What's new:** Property **listLevel** of type **int** </br>
**Description:** Gets or sets the list level of the paragraph. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=paragraph-listLevel)_ </br>

**Resource name:** [paragraph](resources/paragraph.md) </br>
**What's new:** Property **outlineLevel** of type **int** </br>
**Description:** Gets or sets the outline level for the paragraph. </br>
**Available in requirement set:** WordApiDesktop, 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=paragraph-outlineLevel)_ </br>

**Resource name:** [paragraph](resources/paragraph.md) </br>
**What's new:** Property **tableNestingLevel** of type **int** </br>
**Description:** Gets the level of the paragraph's table. It returns 0 if the paragraph is not in a table. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=paragraph-tableNestingLevel)_ </br>

**Resource name:** [paragraph](resources/paragraph.md) </br>
**What's new:** Relationship **list** of type **[List](list.md)** </br>
**Description:** Gets the List to which this paragraph belongs. Returns null if the paragraph is not in a list. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=paragraph-list)_ </br>

**Resource name:** [paragraph](resources/paragraph.md) </br>
**What's new:** Relationship **next** of type **[Paragraph](paragraph.md)** </br>
**Description:** Gets the next paragraph. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=paragraph-next)_ </br>

**Resource name:** [paragraph](resources/paragraph.md) </br>
**What's new:** Relationship **parentBody** of type **[Body](body.md)** </br>
**Description:** Gets the parent body of the paragraph. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=paragraph-parentBody)_ </br>

**Resource name:** [paragraph](resources/paragraph.md) </br>
**What's new:** Relationship **parentTable** of type **[Table](table.md)** </br>
**Description:** Gets the table that contains the paragraph. Returns null if it is not contained in a table. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=paragraph-parentTable)_ </br>

**Resource name:** [paragraph](resources/paragraph.md) </br>
**What's new:** Relationship **parentTableCell** of type **[TableCell](tablecell.md)** </br>
**Description:** Gets the table cell that contains the paragraph. Returns null if it is not contained in a table cell. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=paragraph-parentTableCell)_ </br>

**Resource name:** [paragraph](resources/paragraph.md) </br>
**What's new:** Relationship **previous** of type **[Paragraph](paragraph.md)** </br>
**Description:** Gets the previous paragraph. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=paragraph-previous)_ </br>

**Resource name:** [resources/paragraph](paragraph.md) </br>
**What's new:** Method **[getRange(rangeLocation: RangeLocation)](resources/paragraph.md#getrangerangelocation-rangelocation)** returning **[Range](resources/range.md)** </br>
**Description:** Gets the whole paragraph, or the starting or ending point of the paragraph, as a range. </br>
**Available in requirement set:** 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=paragraph-getRange)_ </br>

**Resource name:** [resources/paragraph](paragraph.md) </br>
**What's new:** Method **[getTextRanges(punctuationMarks: string[], trimSpacing: bool)](resources/paragraph.md#gettextrangespunctuationmarks-string-trimspacing-bool)** returning **[RangeCollection](resources/rangecollection.md)** </br>
**Description:** Gets the text ranges in the paragraph by using punctuation marks andor space character. </br>
**Available in requirement set:** 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=paragraph-getTextRanges)_ </br>

**Resource name:** [resources/paragraph](paragraph.md) </br>
**What's new:** Method **[insertTable(rowCount: number, columnCount: number, insertLocation: InsertLocation, values: string[][])](resources/paragraph.md#inserttablerowcount-number-columncount-number-insertlocation-insertlocation-values-string)** returning **[Table](resources/table.md)** </br>
**Description:** Inserts a table with the specified number of rows and columns. The insertLocation value can be 'Before' or 'After'. </br>
**Available in requirement set:** 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=paragraph-insertTable)_ </br>

**Resource name:** [resources/paragraph](paragraph.md) </br>
**What's new:** Method **[split(delimiters: string[], trimDelimiters: bool, trimSpacing: bool)](resources/paragraph.md#splitdelimiters-string-trimdelimiters-bool-trimspacing-bool)** returning **[RangeCollection](resources/rangecollection.md)** </br>
**Description:** Splits the paragraph into child ranges by using delimiters. </br>
**Available in requirement set:** 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=paragraph-split)_ </br>

**Resource name:** [paragraphCollection](resources/paragraphcollection.md) </br>
**What's new:** Relationship **first** of type **[Paragraph](paragraph.md)** </br>
**Description:** Gets the first paragraph in this collection. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=paragraphCollection-first)_ </br>

**Resource name:** [paragraphCollection](resources/paragraphcollection.md) </br>
**What's new:** Relationship **last** of type **[Paragraph](paragraph.md)** </br>
**Description:** Gets the last paragraph in this collection. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=paragraphCollection-last)_ </br>

**Resource name:** [range](resources/range.md) </br>
**What's new:** Property **hyperlink** of type **string** </br>
**Description:** Gets the first hyperlink in the range, or sets a hyperlink on the range. Existing hyperlinks in this range are deleted when you set a new hyperlink. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=range-hyperlink)_ </br>

**Resource name:** [range](resources/range.md) </br>
**What's new:** Property **isEmpty** of type **bool** </br>
**Description:** Checks whether the range length is zero. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=range-isEmpty)_ </br>

**Resource name:** [range](resources/range.md) </br>
**What's new:** Relationship **lists** of type **[ListCollection](listcollection.md)** </br>
**Description:** Gets the collection of list objects in the range. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=range-lists)_ </br>

**Resource name:** [range](resources/range.md) </br>
**What's new:** Relationship **parentBody** of type **[Body](body.md)** </br>
**Description:** Gets the parent body of the range. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=range-parentBody)_ </br>

**Resource name:** [range](resources/range.md) </br>
**What's new:** Relationship **parentTable** of type **[Table](table.md)** </br>
**Description:** Gets the table that contains the range. Returns null if it is not contained in a table. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=range-parentTable)_ </br>

**Resource name:** [range](resources/range.md) </br>
**What's new:** Relationship **parentTableCell** of type **[TableCell](tablecell.md)** </br>
**Description:** Gets the table cell that contains the range. Returns null if it is not contained in a table cell. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=range-parentTableCell)_ </br>

**Resource name:** [range](resources/range.md) </br>
**What's new:** Relationship **tables** of type **[TableCollection](tablecollection.md)** </br>
**Description:** Gets the collection of table objects in the range. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=range-tables)_ </br>

**Resource name:** [resources/range](range.md) </br>
**What's new:** Method **[compareLocationWith(range: Range)](resources/range.md#comparelocationwithrange-range)** returning **[LocationRelation](resources/locationrelation.md)** </br>
**Description:** Compares this range's location with another range's location. </br>
**Available in requirement set:** 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=range-compareLocationWith)_ </br>

**Resource name:** [resources/range](range.md) </br>
**What's new:** Method **[expandTo(range: Range)](resources/range.md#expandtorange-range)** returning **void** </br>
**Description:** Expands the range in either direction to cover another range. </br>
**Available in requirement set:** 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=range-expandTo)_ </br>

**Resource name:** [resources/range](range.md) </br>
**What's new:** Method **[getHyperlinkRanges()](resources/range.md#gethyperlinkranges)** returning **[RangeCollection](resources/rangecollection.md)** </br>
**Description:** Gets hyperlink child ranges within the range. </br>
**Available in requirement set:** 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=range-getHyperlinkRanges)_ </br>

**Resource name:** [resources/range](range.md) </br>
**What's new:** Method **[getNextTextRange(punctuationMarks: string[], trimSpacing: bool)](resources/range.md#getnexttextrangepunctuationmarks-string-trimspacing-bool)** returning **[Range](resources/range.md)** </br>
**Description:** Gets the next text range by using punctuation marks andor space character. </br>
**Available in requirement set:** 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=range-getNextTextRange)_ </br>

**Resource name:** [resources/range](range.md) </br>
**What's new:** Method **[getRange(rangeLocation: RangeLocation)](resources/range.md#getrangerangelocation-rangelocation)** returning **[Range](resources/range.md)** </br>
**Description:** Clones the range, or gets the starting or ending point of the range as a new range. </br>
**Available in requirement set:** 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=range-getRange)_ </br>

**Resource name:** [resources/range](range.md) </br>
**What's new:** Method **[getTextRanges(punctuationMarks: string[], trimSpacing: bool)](resources/range.md#gettextrangespunctuationmarks-string-trimspacing-bool)** returning **[RangeCollection](resources/rangecollection.md)** </br>
**Description:** Gets the text child ranges in the range by using punctuation marks andor space character. </br>
**Available in requirement set:** 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=range-getTextRanges)_ </br>

**Resource name:** [resources/range](range.md) </br>
**What's new:** Method **[insertTable(rowCount: number, columnCount: number, insertLocation: InsertLocation, values: string[][])](resources/range.md#inserttablerowcount-number-columncount-number-insertlocation-insertlocation-values-string)** returning **[Table](resources/table.md)** </br>
**Description:** Inserts a table with the specified number of rows and columns. The insertLocation value can be 'Before' or 'After'. </br>
**Available in requirement set:** 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=range-insertTable)_ </br>

**Resource name:** [resources/range](range.md) </br>
**What's new:** Method **[intersectWith(range: Range)](resources/range.md#intersectwithrange-range)** returning **void** </br>
**Description:** Shrinks the range to the intersection of the range with another range. </br>
**Available in requirement set:** 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=range-intersectWith)_ </br>

**Resource name:** [resources/range](range.md) </br>
**What's new:** Method **[split(delimiters: string[], multiParagraphs: bool, trimDelimiters: bool, trimSpacing: bool)](resources/range.md#splitdelimiters-string-multiparagraphs-bool-trimdelimiters-bool-trimspacing-bool)** returning **[RangeCollection](resources/rangecollection.md)** </br>
**Description:** Splits the range into child ranges by using delimiters. </br>
**Available in requirement set:** 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=range-split)_ </br>

**Resource name:** [rangeCollection](resources/rangecollection.md) </br>
**What's new:** Property **items** of type **[Range[]](range.md)** </br>
**Description:** A collection of range objects. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=rangeCollection-items)_ </br>

**Resource name:** [rangeCollection](resources/rangecollection.md) </br>
**What's new:** Relationship **first** of type **[Range](range.md)** </br>
**Description:** Gets the first range in this collection. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=rangeCollection-first)_ </br>

**Resource name:** [resources/rangeCollection](rangecollection.md) </br>
**What's new:** Method **[getItem(index: number)](resources/rangecollection.md#getitemindex-number)** returning **[Range](resources/range.md)** </br>
**Description:** Gets a range object by its index in the collection. </br>
**Available in requirement set:** 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=rangeCollection-getItem)_ </br>

**Resource name:** [searchResultCollection](resources/searchresultcollection.md) </br>
**What's new:** Relationship **first** of type **[Range](range.md)** </br>
**Description:** Gets the first searched result in this collection. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=searchResultCollection-first)_ </br>

**Resource name:** [section](resources/section.md) </br>
**What's new:** Relationship **next** of type **[Section](section.md)** </br>
**Description:** Gets the next section. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=section-next)_ </br>

**Resource name:** [sectionCollection](resources/sectioncollection.md) </br>
**What's new:** Relationship **first** of type **[Section](section.md)** </br>
**Description:** Gets the first section in this collection. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=sectionCollection-first)_ </br>

**Resource name:** [table](resources/table.md) </br>
**What's new:** Property **headerRowCount** of type **int** </br>
**Description:** Gets and sets the number of header rows. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-headerRowCount)_ </br>

**Resource name:** [table](resources/table.md) </br>
**What's new:** Property **isUniform** of type **bool** </br>
**Description:** Indicates whether all of the table rows are uniform. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-isUniform)_ </br>

**Resource name:** [table](resources/table.md) </br>
**What's new:** Property **nestingLevel** of type **int** </br>
**Description:** Gets the nesting level of the table. Top-level tables have level 1. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-nestingLevel)_ </br>

**Resource name:** [table](resources/table.md) </br>
**What's new:** Property **rowCount** of type **int** </br>
**Description:** Gets the number of rows in the table. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-rowCount)_ </br>

**Resource name:** [table](resources/table.md) </br>
**What's new:** Property **shadingColor** of type **string** </br>
**Description:** Gets and sets the shading color. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-shadingColor)_ </br>

**Resource name:** [table](resources/table.md) </br>
**What's new:** Property **style** of type **string** </br>
**Description:** Gets and sets the name of the table style. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-style)_ </br>

**Resource name:** [table](resources/table.md) </br>
**What's new:** Property **styleBandedColumns** of type **bool** </br>
**Description:** Gets and sets whether the table has banded columns. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-styleBandedColumns)_ </br>

**Resource name:** [table](resources/table.md) </br>
**What's new:** Property **styleBandedRows** of type **bool** </br>
**Description:** Gets and sets whether the table has banded rows. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-styleBandedRows)_ </br>

**Resource name:** [table](resources/table.md) </br>
**What's new:** Property **styleFirstColumn** of type **bool** </br>
**Description:** Gets and sets whether the table has a first column with a special style. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-styleFirstColumn)_ </br>

**Resource name:** [table](resources/table.md) </br>
**What's new:** Property **styleLastColumn** of type **bool** </br>
**Description:** Gets and sets whether the table has a last column with a special style. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-styleLastColumn)_ </br>

**Resource name:** [table](resources/table.md) </br>
**What's new:** Property **styleTotalRow** of type **bool** </br>
**Description:** Gets and sets whether the table has a total (last) row with a special style. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-styleTotalRow)_ </br>

**Resource name:** [table](resources/table.md) </br>
**What's new:** Property **values** of type **string** </br>
**Description:** Gets and sets the text values in the table, as a 2D Javascript array. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-values)_ </br>

**Resource name:** [table](resources/table.md) </br>
**What's new:** Relationship **cellPaddingBottom** of type **[float](float.md)** </br>
**Description:** Gets and sets the default bottom cell padding in points. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-cellPaddingBottom)_ </br>

**Resource name:** [table](resources/table.md) </br>
**What's new:** Relationship **cellPaddingLeft** of type **[float](float.md)** </br>
**Description:** Gets and sets the default left cell padding in points. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-cellPaddingLeft)_ </br>

**Resource name:** [table](resources/table.md) </br>
**What's new:** Relationship **cellPaddingRight** of type **[float](float.md)** </br>
**Description:** Gets and sets the default right cell padding in points. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-cellPaddingRight)_ </br>

**Resource name:** [table](resources/table.md) </br>
**What's new:** Relationship **cellPaddingTop** of type **[float](float.md)** </br>
**Description:** Gets and sets the default top cell padding in points. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-cellPaddingTop)_ </br>

**Resource name:** [table](resources/table.md) </br>
**What's new:** Relationship **font** of type **[Font](font.md)** </br>
**Description:** Gets the font. Use this to get and set font name, size, color, and other properties. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-font)_ </br>

**Resource name:** [table](resources/table.md) </br>
**What's new:** Relationship **height** of type **[float](float.md)** </br>
**Description:** Gets the height of the table in points. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-height)_ </br>

**Resource name:** [table](resources/table.md) </br>
**What's new:** Relationship **next** of type **[Table](table.md)** </br>
**Description:** Gets the next table. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-next)_ </br>

**Resource name:** [table](resources/table.md) </br>
**What's new:** Relationship **paragraphAfter** of type **[Paragraph](paragraph.md)** </br>
**Description:** Gets the paragraph after the table. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-paragraphAfter)_ </br>

**Resource name:** [table](resources/table.md) </br>
**What's new:** Relationship **paragraphBefore** of type **[Paragraph](paragraph.md)** </br>
**Description:** Gets the paragraph before the table. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-paragraphBefore)_ </br>

**Resource name:** [table](resources/table.md) </br>
**What's new:** Relationship **parentContentControl** of type **[ContentControl](contentcontrol.md)** </br>
**Description:** Gets the content control that contains the table. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-parentContentControl)_ </br>

**Resource name:** [table](resources/table.md) </br>
**What's new:** Relationship **parentTable** of type **[Table](table.md)** </br>
**Description:** Gets the table that contains this table. Returns null if it is not contained in a table. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-parentTable)_ </br>

**Resource name:** [table](resources/table.md) </br>
**What's new:** Relationship **parentTableCell** of type **[TableCell](tablecell.md)** </br>
**Description:** Gets the table cell that contains this table. Returns null if it is not contained in a table cell. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-parentTableCell)_ </br>

**Resource name:** [table](resources/table.md) </br>
**What's new:** Relationship **rows** of type **[TableRowCollection](tablerowcollection.md)** </br>
**Description:** Gets all of the table rows. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-rows)_ </br>

**Resource name:** [table](resources/table.md) </br>
**What's new:** Relationship **tables** of type **[TableCollection](tablecollection.md)** </br>
**Description:** Gets the child tables nested one level deeper. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-tables)_ </br>

**Resource name:** [table](resources/table.md) </br>
**What's new:** Relationship **verticalAlignment** of type **[VerticalAlignment](verticalalignment.md)** </br>
**Description:** Gets and sets the vertical alignment of every cell in the table. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-verticalAlignment)_ </br>

**Resource name:** [table](resources/table.md) </br>
**What's new:** Relationship **width** of type **[float](float.md)** </br>
**Description:** Gets and sets the width of the table in points. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-width)_ </br>

**Resource name:** [resources/table](table.md) </br>
**What's new:** Method **[addColumns(insertLocation: InsertLocation, columnCount: number, values: string[][])](resources/table.md#addcolumnsinsertlocation-insertlocation-columncount-number-values-string)** returning **void** </br>
**Description:** Adds columns to the start or end of the table, using the first or last existing column as a template. This is applicable to uniform tables. The string values, if specified, are set in the newly inserted rows. </br>
**Available in requirement set:** 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-addColumns)_ </br>

**Resource name:** [resources/table](table.md) </br>
**What's new:** Method **[addRows(insertLocation: InsertLocation, rowCount: number, values: string[][])](resources/table.md#addrowsinsertlocation-insertlocation-rowcount-number-values-string)** returning **void** </br>
**Description:** Adds rows to the start or end of the table, using the first or last existing row as a template. The string values, if specified, are set in the newly inserted rows. </br>
**Available in requirement set:** 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-addRows)_ </br>

**Resource name:** [resources/table](table.md) </br>
**What's new:** Method **[autoFitContents()](resources/table.md#autofitcontents)** returning **void** </br>
**Description:** Autofits the table columns to the width of their contents. </br>
**Available in requirement set:** 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-autoFitContents)_ </br>

**Resource name:** [resources/table](table.md) </br>
**What's new:** Method **[autoFitWindow()](resources/table.md#autofitwindow)** returning **void** </br>
**Description:** Autofits the table columns to the width of the window. </br>
**Available in requirement set:** 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-autoFitWindow)_ </br>

**Resource name:** [resources/table](table.md) </br>
**What's new:** Method **[clear()](resources/table.md#clear)** returning **void** </br>
**Description:** Clears the contents of the table. </br>
**Available in requirement set:** 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-clear)_ </br>

**Resource name:** [resources/table](table.md) </br>
**What's new:** Method **[delete()](resources/table.md#delete)** returning **void** </br>
**Description:** Deletes the entire table. </br>
**Available in requirement set:** 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-delete)_ </br>

**Resource name:** [resources/table](table.md) </br>
**What's new:** Method **[deleteColumns(columnIndex: number, columnCount: number)](resources/table.md#deletecolumnscolumnindex-number-columncount-number)** returning **void** </br>
**Description:** Deletes specific columns. This is applicable to uniform tables. </br>
**Available in requirement set:** 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-deleteColumns)_ </br>

**Resource name:** [resources/table](table.md) </br>
**What's new:** Method **[deleteRows(rowIndex: number, rowCount: number)](resources/table.md#deleterowsrowindex-number-rowcount-number)** returning **void** </br>
**Description:** Deletes specific rows. </br>
**Available in requirement set:** 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-deleteRows)_ </br>

**Resource name:** [resources/table](table.md) </br>
**What's new:** Method **[distributeColumns()](resources/table.md#distributecolumns)** returning **void** </br>
**Description:** Distributes the column widths evenly. </br>
**Available in requirement set:** 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-distributeColumns)_ </br>

**Resource name:** [resources/table](table.md) </br>
**What's new:** Method **[distributeRows()](resources/table.md#distributerows)** returning **void** </br>
**Description:** Distributes the row heights evenly. </br>
**Available in requirement set:** 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-distributeRows)_ </br>

**Resource name:** [resources/table](table.md) </br>
**What's new:** Method **[getBorderStyle(borderLocation: BorderLocation)](resources/table.md#getborderstyleborderlocation-borderlocation)** returning **[TableBorderStyle](resources/tableborderstyle.md)** </br>
**Description:** Gets the border style for the specified border. </br>
**Available in requirement set:** 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-getBorderStyle)_ </br>

**Resource name:** [resources/table](table.md) </br>
**What's new:** Method **[getCell(rowIndex: number, cellIndex: number)](resources/table.md#getcellrowindex-number-cellindex-number)** returning **[TableCell](resources/tablecell.md)** </br>
**Description:** Gets the table cell at a specified row and column. </br>
**Available in requirement set:** 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-getCell)_ </br>

**Resource name:** [resources/table](table.md) </br>
**What's new:** Method **[getRange(rangeLocation: RangeLocation)](resources/table.md#getrangerangelocation-rangelocation)** returning **[Range](resources/range.md)** </br>
**Description:** Gets the range that contains this table, or the range at the start or end of the table. </br>
**Available in requirement set:** 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-getRange)_ </br>

**Resource name:** [resources/table](table.md) </br>
**What's new:** Method **[insertContentControl()](resources/table.md#insertcontentcontrol)** returning **[ContentControl](resources/contentcontrol.md)** </br>
**Description:** Inserts a content control on the table. </br>
**Available in requirement set:** 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-insertContentControl)_ </br>

**Resource name:** [resources/table](table.md) </br>
**What's new:** Method **[insertParagraph(paragraphText: string, insertLocation: InsertLocation)](resources/table.md#insertparagraphparagraphtext-string-insertlocation-insertlocation)** returning **[Paragraph](resources/paragraph.md)** </br>
**Description:** Inserts a paragraph at the specified location. The insertLocation value can be 'Before' or 'After'. </br>
**Available in requirement set:** 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-insertParagraph)_ </br>

**Resource name:** [resources/table](table.md) </br>
**What's new:** Method **[insertTable(rowCount: number, columnCount: number, insertLocation: InsertLocation, values: string[][])](resources/table.md#inserttablerowcount-number-columncount-number-insertlocation-insertlocation-values-string)** returning **[Table](resources/table.md)** </br>
**Description:** Inserts a table with the specified number of rows and columns. The insertLocation value can be 'Before' or 'After'. </br>
**Available in requirement set:** 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-insertTable)_ </br>

**Resource name:** [resources/table](table.md) </br>
**What's new:** Method **[mergeCells(topRow: number, firstCell: number, bottomRow: number, lastCell: number)](resources/table.md#mergecellstoprow-number-firstcell-number-bottomrow-number-lastcell-number)** returning **[TableCell](resources/tablecell.md)** </br>
**Description:** Merges the cells bounded inclusively by a first and last cell. </br>
**Available in requirement set:** WordApiDesktop, 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-mergeCells)_ </br>

**Resource name:** [resources/table](table.md) </br>
**What's new:** Method **[search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)](resources/table.md#searchsearchtext-string-searchoptions-paramtypestrings.searchoptions)** returning **[SearchResultCollection](resources/searchresultcollection.md)** </br>
**Description:** Performs a search with the specified searchOptions on the scope of the table object. The search results are a collection of range objects. </br>
**Available in requirement set:** 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-search)_ </br>

**Resource name:** [resources/table](table.md) </br>
**What's new:** Method **[select(selectionMode: SelectionMode)](resources/table.md#selectselectionmode-selectionmode)** returning **void** </br>
**Description:** Selects the table, or the position at the start or end of the table, and navigates the Word UI to it. </br>
**Available in requirement set:** 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-select)_ </br>

**Resource name:** [tableBorderStyle](resources/tableborderstyle.md) </br>
**What's new:** Property **color** of type **string** </br>
**Description:** Gets or sets the table border color, as a hex value or name. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableBorderStyle-color)_ </br>

**Resource name:** [tableBorderStyle](resources/tableborderstyle.md) </br>
**What's new:** Relationship **type** of type **[BorderType](bordertype.md)** </br>
**Description:** Gets or sets the type of the table border style. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableBorderStyle-type)_ </br>

**Resource name:** [tableBorderStyle](resources/tableborderstyle.md) </br>
**What's new:** Relationship **width** of type **[float](float.md)** </br>
**Description:** Gets or sets the width, in points, of the table border style. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableBorderStyle-width)_ </br>

**Resource name:** [tableCell](resources/tablecell.md) </br>
**What's new:** Property **cellIndex** of type **int** </br>
**Description:** Gets the index of the cell in its row. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableCell-cellIndex)_ </br>

**Resource name:** [tableCell](resources/tablecell.md) </br>
**What's new:** Property **rowIndex** of type **int** </br>
**Description:** Gets the index of the cell's row in the table. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableCell-rowIndex)_ </br>

**Resource name:** [tableCell](resources/tablecell.md) </br>
**What's new:** Property **shadingColor** of type **string** </br>
**Description:** Gets or sets the shading color of the cell. Color is specified in "#RRGGBB" format or by using the color name. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableCell-shadingColor)_ </br>

**Resource name:** [tableCell](resources/tablecell.md) </br>
**What's new:** Property **value** of type **string** </br>
**Description:** Gets and sets the text of the cell. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableCell-value)_ </br>

**Resource name:** [tableCell](resources/tablecell.md) </br>
**What's new:** Relationship **body** of type **[Body](body.md)** </br>
**Description:** Gets the body object of the cell. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableCell-body)_ </br>

**Resource name:** [tableCell](resources/tablecell.md) </br>
**What's new:** Relationship **cellPaddingBottom** of type **[float](float.md)** </br>
**Description:** Gets and sets the bottom padding of the cell in points. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableCell-cellPaddingBottom)_ </br>

**Resource name:** [tableCell](resources/tablecell.md) </br>
**What's new:** Relationship **cellPaddingLeft** of type **[float](float.md)** </br>
**Description:** Gets and sets the left padding of the cell in points. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableCell-cellPaddingLeft)_ </br>

**Resource name:** [tableCell](resources/tablecell.md) </br>
**What's new:** Relationship **cellPaddingRight** of type **[float](float.md)** </br>
**Description:** Gets and sets the right padding of the cell in points. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableCell-cellPaddingRight)_ </br>

**Resource name:** [tableCell](resources/tablecell.md) </br>
**What's new:** Relationship **cellPaddingTop** of type **[float](float.md)** </br>
**Description:** Gets and sets the top padding of the cell in points. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableCell-cellPaddingTop)_ </br>

**Resource name:** [tableCell](resources/tablecell.md) </br>
**What's new:** Relationship **columnWidth** of type **[float](float.md)** </br>
**Description:** Gets and sets the width of the cell's column in points. This is applicable to uniform tables. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableCell-columnWidth)_ </br>

**Resource name:** [tableCell](resources/tablecell.md) </br>
**What's new:** Relationship **next** of type **[TableCell](tablecell.md)** </br>
**Description:** Gets the next cell. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableCell-next)_ </br>

**Resource name:** [tableCell](resources/tablecell.md) </br>
**What's new:** Relationship **parentRow** of type **[TableRow](tablerow.md)** </br>
**Description:** Gets the parent row of the cell. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableCell-parentRow)_ </br>

**Resource name:** [tableCell](resources/tablecell.md) </br>
**What's new:** Relationship **parentTable** of type **[Table](table.md)** </br>
**Description:** Gets the parent table of the cell. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableCell-parentTable)_ </br>

**Resource name:** [tableCell](resources/tablecell.md) </br>
**What's new:** Relationship **verticalAlignment** of type **[VerticalAlignment](verticalalignment.md)** </br>
**Description:** Gets and sets the vertical alignment of the cell. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableCell-verticalAlignment)_ </br>

**Resource name:** [tableCell](resources/tablecell.md) </br>
**What's new:** Relationship **width** of type **[float](float.md)** </br>
**Description:** Gets the width of the cell in points. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableCell-width)_ </br>

**Resource name:** [resources/tableCell](tablecell.md) </br>
**What's new:** Method **[deleteColumn()](resources/tablecell.md#deletecolumn)** returning **void** </br>
**Description:** Deletes the column containing this cell. This is applicable to uniform tables. </br>
**Available in requirement set:** 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableCell-deleteColumn)_ </br>

**Resource name:** [resources/tableCell](tablecell.md) </br>
**What's new:** Method **[deleteRow()](resources/tablecell.md#deleterow)** returning **void** </br>
**Description:** Deletes the row containing this cell. </br>
**Available in requirement set:** 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableCell-deleteRow)_ </br>

**Resource name:** [resources/tableCell](tablecell.md) </br>
**What's new:** Method **[getBorderStyle(borderLocation: BorderLocation)](resources/tablecell.md#getborderstyleborderlocation-borderlocation)** returning **[TableBorderStyle](resources/tableborderstyle.md)** </br>
**Description:** Gets the border style for the specified border. </br>
**Available in requirement set:** 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableCell-getBorderStyle)_ </br>

**Resource name:** [resources/tableCell](tablecell.md) </br>
**What's new:** Method **[insertColumns(insertLocation: InsertLocation, columnCount: number, values: string[][])](resources/tablecell.md#insertcolumnsinsertlocation-insertlocation-columncount-number-values-string)** returning **void** </br>
**Description:** Adds columns to the left or right of the cell, using the cell's column as a template. This is applicable to uniform tables. The string values, if specified, are set in the newly inserted rows. </br>
**Available in requirement set:** 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableCell-insertColumns)_ </br>

**Resource name:** [resources/tableCell](tablecell.md) </br>
**What's new:** Method **[insertRows(insertLocation: InsertLocation, rowCount: number, values: string[][])](resources/tablecell.md#insertrowsinsertlocation-insertlocation-rowcount-number-values-string)** returning **void** </br>
**Description:** Inserts rows above or below the cell, using the cell's row as a template. The string values, if specified, are set in the newly inserted rows. </br>
**Available in requirement set:** 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableCell-insertRows)_ </br>

**Resource name:** [resources/tableCell](tablecell.md) </br>
**What's new:** Method **[split(rowCount: number, columnCount: number)](resources/tablecell.md#splitrowcount-number-columncount-number)** returning **void** </br>
**Description:** Adds columns to the left or right of the cell, using the existing column as a template. The string values, if specified, are set in the newly inserted rows. </br>
**Available in requirement set:** WordApiDesktop, 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableCell-split)_ </br>

**Resource name:** [tableCellCollection](resources/tablecellcollection.md) </br>
**What's new:** Property **items** of type **[TableCell[]](tablecell.md)** </br>
**Description:** A collection of tableCell objects. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableCellCollection-items)_ </br>

**Resource name:** [tableCellCollection](resources/tablecellcollection.md) </br>
**What's new:** Relationship **first** of type **[TableCell](tablecell.md)** </br>
**Description:** Gets the first table cell in this collection. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableCellCollection-first)_ </br>

**Resource name:** [resources/tableCellCollection](tablecellcollection.md) </br>
**What's new:** Method **[getItem(index: number)](resources/tablecellcollection.md#getitemindex-number)** returning **[TableCell](resources/tablecell.md)** </br>
**Description:** Gets a table cell object by its index in the collection. </br>
**Available in requirement set:** 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableCellCollection-getItem)_ </br>

**Resource name:** [tableCollection](resources/tablecollection.md) </br>
**What's new:** Property **items** of type **[Table[]](table.md)** </br>
**Description:** A collection of table objects. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableCollection-items)_ </br>

**Resource name:** [tableCollection](resources/tablecollection.md) </br>
**What's new:** Relationship **first** of type **[Table](table.md)** </br>
**Description:** Gets the first table in this collection. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableCollection-first)_ </br>

**Resource name:** [resources/tableCollection](tablecollection.md) </br>
**What's new:** Method **[getItem(index: number)](resources/tablecollection.md#getitemindex-number)** returning **[Table](resources/table.md)** </br>
**Description:** Gets a table object by its index in the collection. </br>
**Available in requirement set:** 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableCollection-getItem)_ </br>

**Resource name:** [tableRow](resources/tablerow.md) </br>
**What's new:** Property **cellCount** of type **int** </br>
**Description:** Gets the number of cells in the row. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableRow-cellCount)_ </br>

**Resource name:** [tableRow](resources/tablerow.md) </br>
**What's new:** Property **isHeader** of type **bool** </br>
**Description:** Gets a value that indicates whether the row is a header row. Read-only. To set the number of header rows, use HeaderRowCount on the Table object. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableRow-isHeader)_ </br>

**Resource name:** [tableRow](resources/tablerow.md) </br>
**What's new:** Property **rowIndex** of type **int** </br>
**Description:** Gets the index of the row in its parent table. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableRow-rowIndex)_ </br>

**Resource name:** [tableRow](resources/tablerow.md) </br>
**What's new:** Property **shadingColor** of type **string** </br>
**Description:** Gets and sets the shading color. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableRow-shadingColor)_ </br>

**Resource name:** [tableRow](resources/tablerow.md) </br>
**What's new:** Property **values** of type **string** </br>
**Description:** Gets and sets the text values in the row, as a 1D Javascript array. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableRow-values)_ </br>

**Resource name:** [tableRow](resources/tablerow.md) </br>
**What's new:** Relationship **cellPaddingBottom** of type **[float](float.md)** </br>
**Description:** Gets and sets the default bottom cell padding for the row in points. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableRow-cellPaddingBottom)_ </br>

**Resource name:** [tableRow](resources/tablerow.md) </br>
**What's new:** Relationship **cellPaddingLeft** of type **[float](float.md)** </br>
**Description:** Gets and sets the default left cell padding for the row in points. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableRow-cellPaddingLeft)_ </br>

**Resource name:** [tableRow](resources/tablerow.md) </br>
**What's new:** Relationship **cellPaddingRight** of type **[float](float.md)** </br>
**Description:** Gets and sets the default right cell padding for the row in points. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableRow-cellPaddingRight)_ </br>

**Resource name:** [tableRow](resources/tablerow.md) </br>
**What's new:** Relationship **cellPaddingTop** of type **[float](float.md)** </br>
**Description:** Gets and sets the default top cell padding for the row in points. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableRow-cellPaddingTop)_ </br>

**Resource name:** [tableRow](resources/tablerow.md) </br>
**What's new:** Relationship **cells** of type **[TableCellCollection](tablecellcollection.md)** </br>
**Description:** Gets cells. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableRow-cells)_ </br>

**Resource name:** [tableRow](resources/tablerow.md) </br>
**What's new:** Relationship **font** of type **[Font](font.md)** </br>
**Description:** Gets the font. Use this to get and set font name, size, color, and other properties. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableRow-font)_ </br>

**Resource name:** [tableRow](resources/tablerow.md) </br>
**What's new:** Relationship **next** of type **[TableRow](tablerow.md)** </br>
**Description:** Gets the next row. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableRow-next)_ </br>

**Resource name:** [tableRow](resources/tablerow.md) </br>
**What's new:** Relationship **parentTable** of type **[Table](table.md)** </br>
**Description:** Gets parent table. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableRow-parentTable)_ </br>

**Resource name:** [tableRow](resources/tablerow.md) </br>
**What's new:** Relationship **preferredHeight** of type **[float](float.md)** </br>
**Description:** Gets and sets the preferred height of the row in points. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableRow-preferredHeight)_ </br>

**Resource name:** [tableRow](resources/tablerow.md) </br>
**What's new:** Relationship **verticalAlignment** of type **[VerticalAlignment](verticalalignment.md)** </br>
**Description:** Gets and sets the vertical alignment of the cells in the row. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableRow-verticalAlignment)_ </br>

**Resource name:** [resources/tableRow](tablerow.md) </br>
**What's new:** Method **[clear()](resources/tablerow.md#clear)** returning **void** </br>
**Description:** Clears the contents of the row. </br>
**Available in requirement set:** 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableRow-clear)_ </br>

**Resource name:** [resources/tableRow](tablerow.md) </br>
**What's new:** Method **[delete()](resources/tablerow.md#delete)** returning **void** </br>
**Description:** Deletes the entire row. </br>
**Available in requirement set:** 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableRow-delete)_ </br>

**Resource name:** [resources/tableRow](tablerow.md) </br>
**What's new:** Method **[getBorderStyle(borderLocation: BorderLocation)](resources/tablerow.md#getborderstyleborderlocation-borderlocation)** returning **[TableBorderStyle](resources/tableborderstyle.md)** </br>
**Description:** Gets the border style of the cells in the row. </br>
**Available in requirement set:** 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableRow-getBorderStyle)_ </br>

**Resource name:** [resources/tableRow](tablerow.md) </br>
**What's new:** Method **[insertRows(insertLocation: InsertLocation, rowCount: number, values: string[][])](resources/tablerow.md#insertrowsinsertlocation-insertlocation-rowcount-number-values-string)** returning **void** </br>
**Description:** Inserts rows using this row as a template. If values are specified, inserts the values into the new rows. </br>
**Available in requirement set:** 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableRow-insertRows)_ </br>

**Resource name:** [resources/tableRow](tablerow.md) </br>
**What's new:** Method **[merge()](resources/tablerow.md#merge)** returning **[TableCell](resources/tablecell.md)** </br>
**Description:** Merges the row into one cell. </br>
**Available in requirement set:** WordApiDesktop, 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableRow-merge)_ </br>

**Resource name:** [resources/tableRow](tablerow.md) </br>
**What's new:** Method **[search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)](resources/tablerow.md#searchsearchtext-string-searchoptions-paramtypestrings.searchoptions)** returning **[SearchResultCollection](resources/searchresultcollection.md)** </br>
**Description:** Performs a search with the specified searchOptions on the scope of the row. The search results are a collection of range objects. </br>
**Available in requirement set:** 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableRow-search)_ </br>

**Resource name:** [resources/tableRow](tablerow.md) </br>
**What's new:** Method **[select(selectionMode: SelectionMode)](resources/tablerow.md#selectselectionmode-selectionmode)** returning **void** </br>
**Description:** Selects the row and navigates the Word UI to it. </br>
**Available in requirement set:** 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableRow-select)_ </br>

**Resource name:** [tableRowCollection](resources/tablerowcollection.md) </br>
**What's new:** Property **items** of type **[TableRow[]](tablerow.md)** </br>
**Description:** A collection of tableRow objects. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableRowCollection-items)_ </br>

**Resource name:** [tableRowCollection](resources/tablerowcollection.md) </br>
**What's new:** Relationship **first** of type **[TableRow](tablerow.md)** </br>
**Description:** Gets the first row in this collection. Read-only. </br>
**Available in requirement set:** 1.3 </br>
_[Give Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableRowCollection-first)_ </br>

**Resource name:** [resources/tableRowCollection](tablerowcollection.md) </br>
**What's new:** Method **[getItem(index: number)](resources/tablerowcollection.md#getitemindex-number)** returning **[TableRow](resources/tablerow.md)** </br>
**Description:** Gets a table row object by its index in the collection. </br>
**Available in requirement set:** 1.3 </br>
_[Feedback](https://github.com/OfficeDev/office-js-docs/issues/new?title=tableRowCollection-getItem)_ </br>





## Try it out 

_**Note**: New features cannot be tried out right now._


We've been working on a Snippet Explorer to let you browse through common code snippets and learn how the new APIs work. Give it a try. The code snippets referenced by the Snippet Explorer are available [here](https://officesnippetexplorer.azurewebsites.net/#/snippets/word). 

## Give feedback

We need it, you want to give it. Feedback is much easier to give now that we're on GitHub. Check out the docs and let us know about any [issues](https://github.com/OfficeDev/office-js-docs/issues) you find in those docs by submitting issues directly in this repository. Let us know what you think about the APIs and the general programming experience. 

We suggest you use the tags [office-js] and [word] on StackOverflow for asking questions to the community.
