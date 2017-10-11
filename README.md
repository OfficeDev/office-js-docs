## Excel JavaScript API Open Specification

_Applies to: Excel 2016, Excel Online, Excel for iOS, Excel for Mac_

This page describes the new set of Excel JavaScript APIs that are planned for future releases. It consists of APIs that are in `beta` stage and others that are still in the design phase. Please review and provide your feedback. One of the best ways you can provide input is to open a new issue in GitHub using the links available below.

_**Note**: The following features are still under design and review phase and not yet available as part of the product. The final design is subject to change. Once the feature is made available, the final specification will be published as part of the master repository._

> **Note**: Any API that is listed as **beta** is not ready for production usage. They are made available so that developers can try them out in test and development environments. They are not meant to be used against production/business critical documents. 

> For the requirement sets that are marked as *Beta*, use the specified (or later) version of Office and use the Beta library of the CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js. Entries not listed as *Beta* are generally available and you can continue to use Production CDN library: https://appsforoffice.microsoft.com/lib/1/hosted/office.js

## Future Excel features

* __Chart API additions__: Ease of Series collection traversing, add/remove chart series, set chart series values, read/set the base unit for the specified category axis, category type of chart Axis, display unit and custom display unit of axis, major and minortime unit scale, axis type, clear the border color of a chart element, sets the border formatting of a chart element to a uniform color, chart font, chart point border, add and manage trendline, chart title string, chart line formatting, etc.

* __Data validation__: Setup data validation rules (drop-down buttons) based on pre-defined list, formulas, date, time or custom values, manage input and error messages. 

* __Document properties__: Access built-in properties, create and manage custom document properties. 
* __Styles__: Create custom styles, apply custom and built style.

* __Hyperlinks__: Add and remove hyperlinks on a range. 

* __Workbook name__: Access workbook name.

* __Gridlines__: Turn gridlines on/off. 

* __Tab color__: Set tab color. 

* __Text orientation__: Range text orientation. 

* __Worksheet headings__: Add worksheet headings. 

* __New range functions__: New additional range handling functions. 

* __Freeze rows/columns__: Freezing rows and columns. 

* __More events__: For details, see [Introduction to Excel 1.8 event features](Event_README.md)

## Upcoming Excel 1.8 event features

This section describes our early thoughts about events in the Excel JavaScript API. We'll add more details about event argument design and code examples in the future. For details, see [Introduction to Excel 1.8 Event Features](Event_README.md).

1.	Data change event for Worksheet, Table with coauthoring support.
    This type of event is triggered when a data change happens; for example, when cell values change in the grid.

2.	Selection change event for Worksheet, and Table.
    This type of event is triggered when the user changes the selection, or extends the selected area.

3.	Activate and Deactivate events for Worksheet.
    This type of event is triggered when a worksheet is activated; for example, when a user switches the worksheet.

4.	Add event for Worksheet with coauthoring support.
    This type of event is triggered when a new worksheet is added to a workbook.


## Key upcoming features

|Object| What is new| Description |Requirement set|
|:----|:----|:----|:----|
|[basicDataValidation](reference/excel/basicdatavalidation.md)|_Property_ > formula1|Gets or sets the Formula1, i.e. minimum value or value depending of the operator.|1.7|
|[basicDataValidation](reference/excel/basicdatavalidation.md)|_Property_ > formula2|Gets or sets the Formula1, i.e. maximum value or value depending of the operator.|1.7|
|[basicDataValidation](reference/excel/basicdatavalidation.md)|_Property_ > operator|The operator to use for validating the data. Possible values are: Invalid, Between, NotBetween, EqualTo, NotEqualTo, GreaterThan, LessThan, GreaterThanOrEqualTo, LessThanOrEqualTo.|1.7|
|[customDataValidation](reference/excel/customdatavalidation.md)|_Property_ > formula|Custom data validation formula, it is to create special rules, such as preventing duplicates, or limiting the total in a range of cells.|1.7|
|[customProperty](reference/excel/customproperty.md)|_Property_ > key|Gets the key of the custom property. Read only. Read-only.|1.7|
|[customProperty](reference/excel/customproperty.md)|_Property_ > type|Gets the value type of the custom property. Read only. Read-only. Possible values are: Number, Boolean, Date, String, Float.|1.7|
|[customProperty](reference/excel/customproperty.md)|_Property_ > value|Gets or sets the value of the custom property.|1.7|
|[customProperty](reference/excel/customproperty.md)|_Method_ > [delete()](reference/excel/customproperty.md#delete)|Deletes the custom property.|1.7|
|[customPropertyCollection](reference/excel/custompropertycollection.md)|_Property_ > items|A collection of customProperty objects. Read-only.|1.7|
|[customPropertyCollection](reference/excel/custompropertycollection.md)|_Method_ > [add(key: string, value: object)](reference/excel/custompropertycollection.md#addkey-string-value-object)|Creates a new or sets an existing custom property.|1.7|
|[customPropertyCollection](reference/excel/custompropertycollection.md)|_Method_ > [deleteAll()](reference/excel/custompropertycollection.md#deleteall)|Deletes all custom properties in this collection.|1.7|
|[customPropertyCollection](reference/excel/custompropertycollection.md)|_Method_ > [getCount()](reference/excel/custompropertycollection.md#getcount)|Gets the count of custom properties.|1.7|
|[customPropertyCollection](reference/excel/custompropertycollection.md)|_Method_ > [getItem(key: string)](reference/excel/custompropertycollection.md#getitemkey-string)|Gets a custom property object by its key, which is case-insensitive. Throws if the custom property does not exist.|1.7|
|[customPropertyCollection](reference/excel/custompropertycollection.md)|_Method_ > [getItemOrNullObject(key: string)](reference/excel/custompropertycollection.md#getitemornullobjectkey-string)|Gets a custom property object by its key, which is case-insensitive. Returns a null object if the custom property does not exist.|1.7|
|[dataValidation](reference/excel/datavalidation.md)|_Property_ > ignoreBlanks|Ignore blanks: no data validation will be performed on blank cells, it defaults to true.|1.7|
|[dataValidation](reference/excel/datavalidation.md)|_Property_ > type|Type of the data validation, see Excel.DataValidationType for details. Read-only. Possible values are: Invalid, None, WholeNumber, Decimal, List, Date, Time, TextLength, Inconsistent, MixedCriteria.|1.7|
|[dataValidation](reference/excel/datavalidation.md)|_Relationship_ > errorAlert|Error alert when user enters invalid data.|1.7|
|[dataValidation](reference/excel/datavalidation.md)|_Relationship_ > prompt|Prompt when users select a cell.|1.7|
|[dataValidation](reference/excel/datavalidation.md)|_Relationship_ > rule|Data Validation rule that contains different type of data validation criteria.|1.7|
|[dataValidation](reference/excel/datavalidation.md)|_Method_ > [clear()](reference/excel/datavalidation.md#clear)|Clears the data validation from the current range.|1.7|
|[dataValidationErrorAlert](reference/excel/datavalidationerroralert.md)|_Property_ > message|Represents error alert message.|1.7|
|[dataValidationErrorAlert](reference/excel/datavalidationerroralert.md)|_Property_ > showAlert|It determines show error alert dialog or not when users enter invalid data, it defaults to true.|1.7|
|[dataValidationErrorAlert](reference/excel/datavalidationerroralert.md)|_Property_ > style|Represents Data validation alert type, please see Excel.DataValidationAlertStyle for details. Possible values are: Invalid, Stop, Warning, Information.|1.7|
|[dataValidationErrorAlert](reference/excel/datavalidationerroralert.md)|_Property_ > title|Represents error alert dialog title.|1.7|
|[dataValidationPrompt](reference/excel/datavalidationprompt.md)|_Property_ > message|Represents the message of the prompt.|1.7|
|[dataValidationPrompt](reference/excel/datavalidationprompt.md)|_Property_ > showPrompt|It determines showing the prompt or not when user selects a cell with the data validation.|1.7|
|[dataValidationPrompt](reference/excel/datavalidationprompt.md)|_Property_ > title|Represents the title for the prompt.|1.7|
|[dataValidationRule](reference/excel/datavalidationrule.md)|_Relationship_ > custom|Custom data validation criteria.|1.7|
|[dataValidationRule](reference/excel/datavalidationrule.md)|_Relationship_ > date|Date data validation criteria.|1.7|
|[dataValidationRule](reference/excel/datavalidationrule.md)|_Relationship_ > decimal|Decimal data validation criteria.|1.7|
|[dataValidationRule](reference/excel/datavalidationrule.md)|_Relationship_ > list|List data validation criteria.|1.7|
|[dataValidationRule](reference/excel/datavalidationrule.md)|_Relationship_ > textLength|TextLength data validation criteria.|1.7|
|[dataValidationRule](reference/excel/datavalidationrule.md)|_Relationship_ > time|Time data validation criteria.|1.7|
|[dataValidationRule](reference/excel/datavalidationrule.md)|_Relationship_ > wholeNumber|WholeNumber data validation criteria.|1.7|
|[dateTimeDataValidation](reference/excel/datetimedatavalidation.md)|_Property_ > formula1|Gets or sets the Formula1, i.e. minimum value or value depending of the operator.|1.7|
|[dateTimeDataValidation](reference/excel/datetimedatavalidation.md)|_Property_ > formula2|Gets or sets the Formula2, i.e. maximum value or value depending of the operator.|1.7|
|[dateTimeDataValidation](reference/excel/datetimedatavalidation.md)|_Property_ > operator|The operator to use for validating the data. Possible values are: Invalid, Between, NotBetween, EqualTo, NotEqualTo, GreaterThan, LessThan, GreaterThanOrEqualTo, LessThanOrEqualTo.|1.7|
|[documentProperties](reference/excel/documentproperties.md)|_Property_ > author|Gets or sets the author of the workbook.|1.7|
|[documentProperties](reference/excel/documentproperties.md)|_Property_ > category|Gets or sets the category of the workbook.|1.7|
|[documentProperties](reference/excel/documentproperties.md)|_Property_ > comments|Gets or sets the comments of the workbook.|1.7|
|[documentProperties](reference/excel/documentproperties.md)|_Property_ > company|Gets or sets the company of the workbook.|1.7|
|[documentProperties](reference/excel/documentproperties.md)|_Property_ > keywords|Gets or sets the keywords of the workbook.|1.7|
|[documentProperties](reference/excel/documentproperties.md)|_Property_ > lastAuthor|Gets the last author of the workbook. Read only. Read-only.|1.7|
|[documentProperties](reference/excel/documentproperties.md)|_Property_ > manager|Gets or sets the manager of the workbook.|1.7|
|[documentProperties](reference/excel/documentproperties.md)|_Property_ > revisionNumber|Gets the revision number of the workbook. Read only. Read-only.|1.7|
|[documentProperties](reference/excel/documentproperties.md)|_Property_ > subject|Gets or sets the subject of the workbook.|1.7|
|[documentProperties](reference/excel/documentproperties.md)|_Property_ > title|Gets or sets the title of the workbook.|1.7|
|[documentProperties](reference/excel/documentproperties.md)|_Relationship_ > creationDate|Gets the creation date of the workbook. Read only. Read-only.|1.7|
|[documentProperties](reference/excel/documentproperties.md)|_Relationship_ > custom|Gets the collection of custom properties of the workbook. Read only. Read-only.|1.7|
|[listDataValidation](reference/excel/listdatavalidation.md)|_Property_ > inCellDropDown|Displays the list in cell drop down or not, it defaults to true.|1.7|
|[listDataValidation](reference/excel/listdatavalidation.md)|_Property_ > source|Source of the list for data validation|1.7|
|[namedItem](reference/excel/nameditem.md)|_Property_ > formula|Gets or sets the formula of the named item.  Formula always starts with a '=' sign.|ApiSet.InProgressFeatures.ArrayValues|
|[namedItem](reference/excel/nameditem.md)|_Relationship_ > arrayValues|Returns an object containing values and types of the named item. Read-only.|ApiSet.InProgressFeatures.ArrayValues|
|[namedItemArrayValues](reference/excel/nameditemarrayvalues.md)|_Property_ > types|Represents the types for each item in the named item array Read-only. Possible values are: Unknown, Empty, String, Integer, Double, Boolean, Error.|ApiSet.InProgressFeatures.ArrayValues|
|[namedItemArrayValues](reference/excel/nameditemarrayvalues.md)|_Property_ > values|Represents the values of each item in the named item array. Read-only.|ApiSet.InProgressFeatures.ArrayValues|
|[pivotTable](reference/excel/pivottable.md)|_Property_ > columnGrandTotals|True if the PivotTable report shows grand totals for columns.|TBD|
|[pivotTable](reference/excel/pivottable.md)|_Property_ > rowGrandTotals|True if the PivotTable report shows grand totals for rows.|TBD|
|[pivotTable](reference/excel/pivottable.md)|_Method_ > [rowAxisLayout(RowLayout: string)](reference/excel/pivottable.md#rowaxislayoutrowlayout-string)|This method is used for simultaneously setting layout options for all existing PivotFields.|TBD|
|[pivotTable](reference/excel/pivottable.md)|_Method_ > [subtotalLocation(Location: string)](reference/excel/pivottable.md#subtotallocationlocation-string)|This method changes the subtotal location for all existing PivotFields. Changing the subtotal location has an immediate visual effect only for fields in outline form, but it will be set for fields in tabular form as well.|TBD|
|[range](reference/excel/range.md)|_Property_ > isEntireColumn|Represents if the current range is an entire column. Read-only.|1.7|
|[range](reference/excel/range.md)|_Property_ > isEntireRow|Represents if the current range is an entire row. Read-only.|1.7|
|[range](reference/excel/range.md)|_Property_ > numberFormatLocal|Represents Excel's number format code for the given cell as a string in the language of the user.|1.7|
|[range](reference/excel/range.md)|_Property_ > style|Represents the style of the current range. This return either null or a string.|1.7|
|[range](reference/excel/range.md)|_Relationship_ > dataValidation|Returns a data validation object. Read-only.|1.7|
|[range](reference/excel/range.md)|_Relationship_ > hyperlink|Represents the hyperlink set for the current range.|1.7|
|[range](reference/excel/range.md)|_Method_ > [getAbsoluteResizedRange(numRows: number, numColumns: number)](reference/excel/range.md#getabsoluteresizedrangenumrows-number-numcolumns-number)|Gets a Range object with the same top-left cell as the current Range object, but with the specified numbers of rows and columns.|1.7|
|[range](reference/excel/range.md)|_Method_ > [getImage()](reference/excel/range.md#getimage)|Renders the range as a base64-encoded image.|1.7|
|[range](reference/excel/range.md)|_Method_ > [getSurroundingRegion()](reference/excel/range.md#getsurroundingregion)|Returns a Range object that represents the surrounding region for the top-left cell in this range. A surrounding region is a range bounded by any combination of blank rows and blank columns relative to this range.|1.7|
|[range](reference/excel/range.md)|_Method_ > [showCard()](reference/excel/range.md#showcard)|Displays the card for an active cell if it has rich value content.|TBD|
|[rangeFormat](reference/excel/rangeformat.md)|_Property_ > textOrientation|Gets or sets the text orientation of all the cells within the range.|1.7|
|[rangeFormat](reference/excel/rangeformat.md)|_Property_ > useStandardHeight|Determines if the row height of the Range object equals the standard height of the sheet.|1.7|
|[rangeFormat](reference/excel/rangeformat.md)|_Property_ > useStandardWidth|Determines if the columnwidth of the Range object equals the standard width of the sheet.|1.7|
|[rangeHyperlink](reference/excel/rangehyperlink.md)|_Property_ > address|Represents the url target for the hyperlink.|1.7|
|[rangeHyperlink](reference/excel/rangehyperlink.md)|_Property_ > documentReference|Represents the document reference target for the hyperlink.|1.7|
|[rangeHyperlink](reference/excel/rangehyperlink.md)|_Property_ > screenTip|Represents the string displayed when hovering over the hyperlink.|1.7|
|[rangeHyperlink](reference/excel/rangehyperlink.md)|_Property_ > textToDisplay|Represents the string that is displayed in the top left most cell in the range.|1.7|
|[rangeViewCollection](reference/excel/rangeviewcollection.md)|_Method_ > [getFirst()](reference/excel/rangeviewcollection.md#getfirst)|Gets the first RangeView object in the collection.|TBD|
|[rangeViewCollection](reference/excel/rangeviewcollection.md)|_Method_ > [getLast()](reference/excel/rangeviewcollection.md#getlast)|Gets the last RangeView object in the collection.|TBD|
|[showCardPostProcessAction](reference/excel/showcardpostprocessaction.md)|_Property_ > actionType|Transmits additional data to client side, e.g., worksheetId for TableSelectionChangedEvent. Possible values are: None, RegisterEvent, UnregisterEvent, CustomFunction, ShowCard.|TBD|
|[showCardPostProcessAction](reference/excel/showcardpostprocessaction.md)|_Property_ > column|Transmits additional data to client side, e.g., worksheetId for TableSelectionChangedEvent. Possible values are: None, RegisterEvent, UnregisterEvent, CustomFunction, ShowCard.|TBD|
|[showCardPostProcessAction](reference/excel/showcardpostprocessaction.md)|_Property_ > row|Transmits additional data to client side, e.g., worksheetId for TableSelectionChangedEvent. Possible values are: None, RegisterEvent, UnregisterEvent, CustomFunction, ShowCard.|TBD|
|[style](reference/excel/style.md)|_Property_ > addIndent|Indicates if text is automatically indented when the text alignment in a cell is set to equal distribution.|1.7|
|[style](reference/excel/style.md)|_Property_ > builtIn|Indicates if the style is a built-in style. Read-only.|1.7|
|[style](reference/excel/style.md)|_Property_ > formulaHidden|Indicates if the formula will be hidden when the worksheet is protected.|1.7|
|[style](reference/excel/style.md)|_Property_ > horizontalAlignment|Represents the horizontal alignment for the style. Possible values are: General, Left, Center, Right, Fill, Justify, CenterAcrossSelection, Distributed.|1.7|
|[style](reference/excel/style.md)|_Property_ > includeAlignment|Indicates if the style includes the AddIndent, HorizontalAlignment, VerticalAlignment, WrapText, IndentLevel, and Orientation properties.|1.7|
|[style](reference/excel/style.md)|_Property_ > includeBorder|Indicates if the style includes the Color, ColorIndex, LineStyle, and Weight border properties.|1.7|
|[style](reference/excel/style.md)|_Property_ > includeFont|Indicates if the style includes the Background, Bold, Color, ColorIndex, FontStyle, Italic, Name, Size, Strikethrough, Subscript, Superscript, and Underline font properties.|1.7|
|[style](reference/excel/style.md)|_Property_ > includeNumber|Indicates if the style includes the NumberFormat property.|1.7|
|[style](reference/excel/style.md)|_Property_ > includePatterns|Indicates if the style includes the Color, ColorIndex, InvertIfNegative, Pattern, PatternColor, and PatternColorIndex interior properties.|1.7|
|[style](reference/excel/style.md)|_Property_ > includeProtection|Indicates if the style includes the FormulaHidden and Locked protection properties.|1.7|
|[style](reference/excel/style.md)|_Property_ > indentLevel|An integer from 0 to 15 that indicates the indent level for the style.|1.7|
|[style](reference/excel/style.md)|_Property_ > locked|Indicates if the object is locked when the worksheet is protected.|1.7|
|[style](reference/excel/style.md)|_Property_ > name|The name of the style. Read-only.|1.7|
|[style](reference/excel/style.md)|_Property_ > numberFormat|The format code of the number format for the style.|1.7|
|[style](reference/excel/style.md)|_Property_ > numberFormatLocal|The localized format code of the number format for the style.|1.7|
|[style](reference/excel/style.md)|_Property_ > orientation|The text orientation for the style.|1.7|
|[style](reference/excel/style.md)|_Property_ > readingOrder|The reading order for the style. Possible values are: Context, LeftToRight, RightToLeft.|1.7|
|[style](reference/excel/style.md)|_Property_ > shrinkToFit|Indicates if text automatically shrinks to fit in the available column width.|1.7|
|[style](reference/excel/style.md)|_Property_ > verticalAlignment|Represents the vertical alignment for the style. Possible values are: Top, Center, Bottom, Justify, Distributed.|1.7|
|[style](reference/excel/style.md)|_Property_ > wrapText|Indicates if Microsoft Excel wraps the text in the object.|1.7|
|[style](reference/excel/style.md)|_Relationship_ > borders|A Border collection of four Border objects that represent the style of the four borders. Read-only.|1.7|
|[style](reference/excel/style.md)|_Relationship_ > fill|The Fill of the style. Read-only.|1.7|
|[style](reference/excel/style.md)|_Relationship_ > font|A Font object that represents the font of the style. Read-only.|1.7|
|[style](reference/excel/style.md)|_Method_ > [delete()](reference/excel/style.md#delete)|Deletes this style.|1.7|
|[styleCollection](reference/excel/stylecollection.md)|_Property_ > items|A collection of style objects. Read-only.|1.7|
|[styleCollection](reference/excel/stylecollection.md)|_Method_ > [add(name: string)](reference/excel/stylecollection.md#addname-string)|Adds a new style to the collection.|1.7|
|[styleCollection](reference/excel/stylecollection.md)|_Method_ > [getItem(name: string)](reference/excel/stylecollection.md#getitemname-string)|Gets a style by name.|1.7|
|[tableColumn](reference/excel/tablecolumn.md)|_Method_ > [getNextColumn()](reference/excel/tablecolumn.md#getnextcolumn)|Gets the table column that follows this one. If there are no table columns following this one, this method will throw an error.|TBD|
|[tableColumn](reference/excel/tablecolumn.md)|_Method_ > [getNextColumnOrNullObject()](reference/excel/tablecolumn.md#getnextcolumnornullobject)|Gets the table column that follows this one. If there are no table columns following this one, this method will return a null object.|TBD|
|[tableColumn](reference/excel/tablecolumn.md)|_Method_ > [getPreviousColumn()](reference/excel/tablecolumn.md#getpreviouscolumn)|Gets the table column that precedes this one. If there are no previous table columns, this method will throw an error.|TBD|
|[tableColumn](reference/excel/tablecolumn.md)|_Method_ > [getPreviousColumnOrNullObject()](reference/excel/tablecolumn.md#getpreviouscolumnornullobject)|Gets the table column that precedes this one. If there are no previous table columns, this method will return a null objet.|TBD|
|[tableColumnCollection](reference/excel/tablecolumncollection.md)|_Method_ > [getFirst()](reference/excel/tablecolumncollection.md#getfirst)|Gets the first column in the table.|TBD|
|[tableColumnCollection](reference/excel/tablecolumncollection.md)|_Method_ > [getLast()](reference/excel/tablecolumncollection.md#getlast)|Gets the last column in the table.|TBD|
|[workbook](reference/excel/workbook.md)|_Property_ > name|Gets the workbook name. Read-only.|1.7|
|[workbook](reference/excel/workbook.md)|_Relationship_ > properties|Gets the workbook properties. Read-only.|1.7|
|[workbook](reference/excel/workbook.md)|_Relationship_ > styles|Represents a collection of styles associated with the workbook. Read-only.|1.7|
|[workbook](reference/excel/workbook.md)|_Method_ > [getActiveCell()](reference/excel/workbook.md#getactivecell)|Gets the currently active cell from the workbook.|1.7|
|[workbook](reference/excel/workbook.md)|_Method_ > [getRange(address: string)](reference/excel/workbook.md#getrangeaddress-string)|Gets the range object specified by the address or name.|ApiSet.InProgressFeatures.WorkbookRange|
|[worksheet](reference/excel/worksheet.md)|_Property_ > gridlines|Gets or sets the worksheet's gridlines flag.|1.7|
|[worksheet](reference/excel/worksheet.md)|_Property_ > headings|Gets or sets the worksheet's headings flag.|1.7|
|[worksheet](reference/excel/worksheet.md)|_Property_ > standardHeight|Returns the standard (default) height of all the rows in the worksheet, in points. Read-only.|1.7|
|[worksheet](reference/excel/worksheet.md)|_Property_ > standardWidth|Returns or sets the standard (default) width of all the columns in the worksheet. Readwrite.|1.7|
|[worksheet](reference/excel/worksheet.md)|_Property_ > tabColor|Gets or sets the worksheet tab color.|1.7|
|[worksheet](reference/excel/worksheet.md)|_Relationship_ > freezePanes|Gets an object that can be used to manipulate frozen panes on the worksheet Read-only.|1.7|
|[worksheet](reference/excel/worksheet.md)|_Relationship_ > pageLayout|Gets the PageLayout object of the worksheet. Read-only.|1.7|
|[worksheet](reference/excel/worksheet.md)|_Method_ > [getRangeByIndexes(startRow: number, startColumn: number, rowCount: number, columnCount: number)](reference/excel/worksheet.md#getrangebyindexesstartrow-number-startcolumn-number-rowcount-number-columncount-number)|Gets the range object beginning at a particular row index and column index, and spanning a certain number of rows and columns.|1.7|
|[worksheetFreezePanes](reference/excel/worksheetfreezepanes.md)|_Method_ > [freezeAt(frozenRange: Range or string)](reference/excel/worksheetfreezepanes.md#freezeatfrozenrange-range-or-string)|Sets the frozen cells in the active worksheet view.|1.7|
|[worksheetFreezePanes](reference/excel/worksheetfreezepanes.md)|_Method_ > [freezeColumns(count: number)](reference/excel/worksheetfreezepanes.md#freezecolumnscount-number)|Freeze the first column(s) of the worksheet in place.|1.7|
|[worksheetFreezePanes](reference/excel/worksheetfreezepanes.md)|_Method_ > [freezeRows(count: number)](reference/excel/worksheetfreezepanes.md#freezerowscount-number)|Freeze the top row(s) of the worksheet in place.|1.7|
|[worksheetFreezePanes](reference/excel/worksheetfreezepanes.md)|_Method_ > [getLocation()](reference/excel/worksheetfreezepanes.md#getlocation)|Gets a range that describes the frozen cells in the active worksheet view.|1.7|
|[worksheetFreezePanes](reference/excel/worksheetfreezepanes.md)|_Method_ > [getLocationOrNullObject()](reference/excel/worksheetfreezepanes.md#getlocationornullobject)|Gets a range that describes the frozen cells in the active worksheet view.|1.7|
|[worksheetFreezePanes](reference/excel/worksheetfreezepanes.md)|_Method_ > [unfreeze()](reference/excel/worksheetfreezepanes.md#unfreeze)|Removes all frozen panes in the worksheet.|1.7|
|[worksheetProtectionOptions](reference/excel/worksheetprotectionoptions.md)|_Property_ > allowEditObjects|Represents the worksheet protection option of allowing editing objects.|TBD|
|[worksheetProtectionOptions](reference/excel/worksheetprotectionoptions.md)|_Property_ > allowEditScenarios|Represents the worksheet protection option of allowing editing scenarios.|TBD|
|[worksheetProtectionOptions](reference/excel/worksheetprotectionoptions.md)|_Property_ > selectionMode|Represents the worksheet protection option of selection mode. Possible values are: Normal, Unlocked, None.|TBD|


## Chart API additions

_Details of the specific APIs are listed below. Please let us know what you think!_

|Object| What is new| Description |Requirement set|
|:----|:----|:----|:----|
|[chart](reference/excel/chart.md)|_Property_ > id|The unique id of chart. Read-only.|TBD|
|[chart](reference/excel/chart.md)|_Property_ > showAllFieldButtons|Returns or sets whether to display all field buttons on a PivotChart. Readwrite|TBD|
|[chartAxes](reference/excel/chartaxes.md)|_Method_ > [getItem(type: string, group: string)](reference/excel/chartaxes.md#getitemtype-string-group-string)|Returns the specific axis identified by type and group.|TBD|
|[chartAxis](reference/excel/chartaxis.md)|_Property_ > axisGroup|Represents the group for the specified axis. Read-only. Possible values are: Primary, Secondary.|TBD|
|[chartAxis](reference/excel/chartaxis.md)|_Property_ > categoryType|Returns or sets the category axis type. Possible values are: Automatic, TextAxis, DateAxis.|TBD|
|[chartAxis](reference/excel/chartaxis.md)|_Property_ > crosses|Represents the specified axis where the other axis crosses. Possible values are: Automatic, Maximum, Minimum, Custom.|TBD|
|[chartAxis](reference/excel/chartaxis.md)|_Property_ > crossesAt|Represents the specified axis where the other axis crosses at. Read Only. Set to this property should use SetCrossesAt(double) method. Read-only.|TBD|
|[chartAxis](reference/excel/chartaxis.md)|_Property_ > customDisplayUnit|Represents the custom axis display unit value. Read Only. To set this property, please use the SetCustomDisplayUnit(double) method. Read-only.|TBD|
|[chartAxis](reference/excel/chartaxis.md)|_Property_ > displayUnit|Represents the axis display unit. Possible values are: None, Hundreds, Thousands, TenThousands, HundredThousands, Millions, TenMillions, HundredMillions, Billions, Trillions, Custom.|TBD|
|[chartAxis](reference/excel/chartaxis.md)|_Property_ > height|Represents the height, in points, of the chart axis. Null if the axis's not visible. Read-only.|TBD|
|[chartAxis](reference/excel/chartaxis.md)|_Property_ > left|Represents the distance, in points, from the left edge of the axis to the left of chart area. Null if the axis's not visible. Read-only.|TBD|
|[chartAxis](reference/excel/chartaxis.md)|_Property_ > logBase|Represents the base of the logarithm when using logarithmic scales.|TBD|
|[chartAxis](reference/excel/chartaxis.md)|_Property_ > reversePlotOrder|Represents whether Microsoft Excel plots data points from last to first.|TBD|
|[chartAxis](reference/excel/chartaxis.md)|_Property_ > scaleType|Represents the value axis scale type. Possible values are: Linear, Logarithmic.|TBD|
|[chartAxis](reference/excel/chartaxis.md)|_Property_ > showDisplayUnitLabel|Represents whether the axis display unit label is visible.|TBD|
|[chartAxis](reference/excel/chartaxis.md)|_Property_ > top|Represents the distance, in points, from the top edge of the axis to the top of chart area. Null if the axis's not visible. Read-only.|TBD|
|[chartAxis](reference/excel/chartaxis.md)|_Property_ > type|Represents the axis type. Read-only. Possible values are: Invalid, Category, Value, Series.|TBD|
|[chartAxis](reference/excel/chartaxis.md)|_Property_ > visible|A boolean value represents the visibility of the axis.|TBD|
|[chartAxis](reference/excel/chartaxis.md)|_Property_ > width|Represents the width, in points, of the chart axis. Null if the axis's not visible. Read-only.|TBD|
|[chartAxis](reference/excel/chartaxis.md)|_Relationship_ > baseTimeUnit|Returns or sets the base unit for the specified category axis.|TBD|
|[chartAxis](reference/excel/chartaxis.md)|_Relationship_ > majorTimeUnitScale|Returns or sets the major unit scale value for the category axis when the CategoryType property is set to TimeScale.|TBD|
|[chartAxis](reference/excel/chartaxis.md)|_Relationship_ > minorTimeUnitScale|Returns or sets the minor unit scale value for the category axis when the CategoryType property is set to TimeScale.|TBD|
|[chartAxis](reference/excel/chartaxis.md)|_Method_ > [setCategoryNames(sourceData: object)](reference/excel/chartaxis.md#setcategorynamessourcedata-object)|Sets all the category names for the specified axis.|TBD|
|[chartAxis](reference/excel/chartaxis.md)|_Method_ > [setCrossesAt(value: double)](reference/excel/chartaxis.md#setcrossesatvalue-double)|Set the specified axis where the other axis crosses at.|TBD|
|[chartAxis](reference/excel/chartaxis.md)|_Method_ > [setCustomDisplayUnit(value: double)](reference/excel/chartaxis.md#setcustomdisplayunitvalue-double)|Sets the axis display unit to a custom value.|TBD|
|[chartDataLabel](reference/excel/chartdatalabel.md)|_Property_ > position|DataLabelPosition value that represents the position of the data label. Possible values are: None, Center, InsideEnd, InsideBase, OutsideEnd, Left, Right, Top, Bottom, BestFit, Callout.|TBD|
|[chartFormatString](reference/excel/chartformatstring.md)|_Relationship_ > font|Represents the font attributes, such as font name, font size, color, etc. of chart characters object. Read-only.|TBD|
|[chartLegend](reference/excel/chartlegend.md)|_Property_ > height|Represents the height of the legend on the chart.|TBD|
|[chartLegend](reference/excel/chartlegend.md)|_Property_ > left|Represents the left of a chart legend.|TBD|
|[chartLegend](reference/excel/chartlegend.md)|_Property_ > showShadow|Represents if the legend has shadow on the chart.|TBD|
|[chartLegend](reference/excel/chartlegend.md)|_Property_ > top|Represents the top of a chart legend.|TBD|
|[chartLegend](reference/excel/chartlegend.md)|_Property_ > width|Represents the width of the legend on the chart.|TBD|
|[chartLegend](reference/excel/chartlegend.md)|_Relationship_ > legendEntries|Represents a collection of legendEntries in the legend. Read-only.|TBD|
|[chartLegendEntry](reference/excel/chartlegendentry.md)|_Property_ > visible|Represents the visible of a chart legend entry.|TBD|
|[chartLegendEntryCollection](reference/excel/chartlegendentrycollection.md)|_Property_ > items|A collection of chartLegendEntry objects. Read-only.|TBD|
|[chartLegendEntryCollection](reference/excel/chartlegendentrycollection.md)|_Method_ > [getCount()](reference/excel/chartlegendentrycollection.md#getcount)|Returns the number of legendEntry in the collection.|TBD|
|[chartLegendEntryCollection](reference/excel/chartlegendentrycollection.md)|_Method_ > [getItemAt(index: number)](reference/excel/chartlegendentrycollection.md#getitematindex-number)|Returns a legendEntry at the given index.|TBD|
|[chartPoint](reference/excel/chartpoint.md)|_Relationship_ > dataLabel|Returns the data label of a chart point. Read-only.|TBD|
|[chartPointsCollection](reference/excel/chartpointscollection.md)|_Method_ > [getFirst()](reference/excel/chartpointscollection.md#getfirst)|Gets the first point in the series.|TBD|
|[chartPointsCollection](reference/excel/chartpointscollection.md)|_Method_ > [getLast()](reference/excel/chartpointscollection.md#getlast)|Gets the last point in the series.|TBD|
|[chartSeries](reference/excel/chartseries.md)|_Property_ > chartType|Represents the chart type of a series. Possible values are: ColumnClustered, ColumnStacked, ColumnStacked100, BarClustered, BarStacked, BarStacked100, LineStacked, LineStacked100, LineMarkers, LineMarkersStacked, LineMarkersStacked100, PieOfPie, etc..|TBD|
|[chartSeries](reference/excel/chartseries.md)|_Property_ > filtered|Boolean value representing if the series is filtered or not. Not applicable for surface charts.|TBD|
|[chartSeries](reference/excel/chartseries.md)|_Property_ > hasDataLabels|Boolean value representing if the series has data labels or not.|TBD|
|[chartSeries](reference/excel/chartseries.md)|_Property_ > markerBackgroundColor|Represents markers background color of a chart series.|TBD|
|[chartSeries](reference/excel/chartseries.md)|_Property_ > markerForegroundColor|Represents markers foreground color of a chart series.|TBD|
|[chartSeries](reference/excel/chartseries.md)|_Property_ > markerSize|Represents marker size of a chart series.|TBD|
|[chartSeries](reference/excel/chartseries.md)|_Property_ > markerStyle|Represents marker style of a chart series. Possible values are: Invalid, Automatic, None, Square, Diamond, Triangle, X, Star, Dot, Dash, Circle, Plus, Picture.|TBD|
|[chartSeries](reference/excel/chartseries.md)|_Property_ > plotOrder|Represents the plot order of a chart series within the chart group.|TBD|
|[chartSeries](reference/excel/chartseries.md)|_Property_ > showShadow|Boolean value representing if the series has shadow or not.|TBD|
|[chartSeries](reference/excel/chartseries.md)|_Property_ > smooth|Boolean value representing if the series is smooth or not. Only for line and scatter charts.|TBD|
|[chartSeries](reference/excel/chartseries.md)|_Relationship_ > trendlines|Represents a collection of trendlines in the series. Read-only.|TBD|
|[chartSeries](reference/excel/chartseries.md)|_Method_ > [delete()](reference/excel/chartseries.md#delete)|Deletes the chart series.|TBD|
|[chartSeries](reference/excel/chartseries.md)|_Method_ > [setBubbleSizes(sourceData: Range)](reference/excel/chartseries.md#setbubblesizessourcedata-range)|Set bubble sizes for a chart series. Only works for bubble charts.|TBD|
|[chartSeries](reference/excel/chartseries.md)|_Method_ > [setValues(sourceData: Range)](reference/excel/chartseries.md#setvaluessourcedata-range)|Set values for a chart series. For scatter chart, it means Y axis values.|TBD|
|[chartSeries](reference/excel/chartseries.md)|_Method_ > [setXAxisValues(sourceData: Range)](reference/excel/chartseries.md#setxaxisvaluessourcedata-range)|Set values of X axis for a chart series. Only works for scatter charts.|TBD|
|[chartSeriesCollection](reference/excel/chartseriescollection.md)|_Method_ > [add(name: string, index: number)](reference/excel/chartseriescollection.md#addname-string-index-number)|Add a new series to the collection.|TBD|
|[chartSeriesCollection](reference/excel/chartseriescollection.md)|_Method_ > [getFirst()](reference/excel/chartseriescollection.md#getfirst)|Gets the first series in the collection.|TBD|
|[chartSeriesCollection](reference/excel/chartseriescollection.md)|_Method_ > [getLast()](reference/excel/chartseriescollection.md#getlast)|Gets the last series in the collection.|TBD|
|[chartTitle](reference/excel/charttitle.md)|_Property_ > height|Returns the height, in points, of the chart title. Read-only. Null if chart title's not visible. Read-only.|TBD|
|[chartTitle](reference/excel/charttitle.md)|_Property_ > horizontalAlignment|Represents the horizontal alignment for chart title. Possible values are: Center, Left, Justify, Distributed, Right.|TBD|
|[chartTitle](reference/excel/charttitle.md)|_Property_ > left|Represents the distance, in points, from the left edge of chart title to the left edge of chart area. Null if chart title's not visible.|TBD|
|[chartTitle](reference/excel/charttitle.md)|_Property_ > position|Represents the position of chart title. Possible values are: Top, Automatic, Bottom, Right, Left.|TBD|
|[chartTitle](reference/excel/charttitle.md)|_Property_ > showShadow|Represents a boolean value that determines if the chart title has a shadow.|TBD|
|[chartTitle](reference/excel/charttitle.md)|_Property_ > textOrientation|Represents the text orientation of chart title. The value should be an integer either from -90 to 90, or 180 for vertically-oriented text.|TBD|
|[chartTitle](reference/excel/charttitle.md)|_Property_ > top|Represents the distance, in points, from the top edge of chart title to the top of chart area. Null if chart title's not visible.|TBD|
|[chartTitle](reference/excel/charttitle.md)|_Property_ > verticalAlignment|Represents the vertical alignment of chart title. Possible values are: Center, Bottom, Top, Justify, Distributed.|TBD|
|[chartTitle](reference/excel/charttitle.md)|_Property_ > width|Returns the width, in points, of the chart title. Read-only. Null if chart title's not visible. Read-only.|TBD|
|[chartTitle](reference/excel/charttitle.md)|_Method_ > [getSubstring(start: number, length: number)](reference/excel/charttitle.md#getsubstringstart-number-length-number)|Get the substring of a chart title. Line break '\n' also counts one charater.|TBD|
|[chartTitle](reference/excel/charttitle.md)|_Method_ > [setFormula(formula: string)](reference/excel/charttitle.md#setformulaformula-string)|Sets a string value that represents the formula of chart title using A1-style notation.|TBD|
|[chartTrendline](reference/excel/charttrendline.md)|_Property_ > backward|Represents the number of periods that the trendline extends backward.|TBD|
|[chartTrendline](reference/excel/charttrendline.md)|_Property_ > displayEquation|True if the equation for the trendline is displayed on the chart.|TBD|
|[chartTrendline](reference/excel/charttrendline.md)|_Property_ > displayRSquared|True if the R-squared for the trendline is displayed on the chart.|TBD|
|[chartTrendline](reference/excel/charttrendline.md)|_Property_ > forward|Represents the number of periods that the trendline extends forward.|TBD|
|[chartTrendline](reference/excel/charttrendline.md)|_Property_ > movingAveragePeriod|Represents the period of a chart trendline, only for trendline with MovingAverage type.|TBD|
|[chartTrendline](reference/excel/charttrendline.md)|_Property_ > polynomialOrder|Represents the order of a chart trendline, only for trendline with Polynomial type.|TBD|
|[chartTrendline](reference/excel/charttrendline.md)|_Property_ > type|Represents the type of a chart trendline. Possible values are: Linear, Exponential, Logarithmic, MovingAverage, Polynomial, Power.|TBD|
|[chartTrendline](reference/excel/charttrendline.md)|_Relationship_ > format|Represents the formatting of a chart trendline. Read-only.|TBD|
|[chartTrendline](reference/excel/charttrendline.md)|_Method_ > [delete()](reference/excel/charttrendline.md#delete)|Delete the trendline object.|TBD|
|[chartTrendlineCollection](reference/excel/charttrendlinecollection.md)|_Property_ > items|A collection of chartTrendline objects. Read-only.|TBD|
|[chartTrendlineCollection](reference/excel/charttrendlinecollection.md)|_Method_ > [add(type: string)](reference/excel/charttrendlinecollection.md#addtype-string)|Adds a new trendline to trendline collection.|TBD|
|[chartTrendlineCollection](reference/excel/charttrendlinecollection.md)|_Method_ > [getCount()](reference/excel/charttrendlinecollection.md#getcount)|Returns the number of trendlines in the collection.|TBD|
|[chartTrendlineCollection](reference/excel/charttrendlinecollection.md)|_Method_ > [getItem(index: number)](reference/excel/charttrendlinecollection.md#getitemindex-number)|Get trendline object by index, which is the insertion order in items array.|TBD|
|[chartTrendlineFormat](reference/excel/charttrendlineformat.md)|_Relationship_ > line|Represents chart line formatting. Read-only.|TBD|




## Give us feedback!

We need it, you want to give it. Feedback is much easier to give now that we're on GitHub. Check out the docs and let us know your thoughts by submitting [issues](https://github.com/OfficeDev/office-js-docs/issues) directly in this repository.

For API support, you can post questions to the community on [StackOverflow](http://stackoverflow.com/questions/tagged/office-js) and tag them with [office-js].
