## Excel JavaScript API Open Specification

_Applies to: Excel 2016, Excel Online, Excel for iOS, Excel for Mac_

This page describes the new set of Excel JavaScript APIs that are being planned for the next releases. It consists of combination of APIs that are in `beta` stage and the ones that are in design phase. Please review and provide your feedback. One of the best ways of providing your input is by opening new issue in GitHub using the links available below.

_**Note**: below listed features are still under design and review phase and hence not yet available as part of the product. The final design is subject to change. Once the feature is made available, the final specification will be published as part of the master repository._

> **Note**: Any API that is listed as **beta** is not ready for production usage. They are made available so that developers can try them out in test and development environments. They are not meant to be used against production/business critical documents. 

> For the requirement sets that are marked as *Beta*, use the specified (or later) version of the Office software and use the Beta library of the CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js. Entries not listed as *Beta* are generally available and you can continue to use Production CDN library: https://appsforoffice.microsoft.com/lib/1/hosted/office.js

## Upcoming Excel 1.7 Release Features

* Chart API additions
* Workbook name
* Gridlines
* Tab color
* Text orientation
* Worksheet headings
* New range functions
* Freeze rows/columns.

## Upcoming Excel 1.8 Event Features

This section describes our early thoughts about events in the Excel JavaScript API. We'll add more details about event argument design and code examples in the future. Please see details in [Introduction to Excel 1.8 Event Features](Event_README.md)

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
|[basicDataValidation](reference/excel/basicdatavalidation.md)|_Property_ > formula1|Gets or sets the Formula1, i.e. minimum value or value depending of the operator.|1.7D|
|[basicDataValidation](reference/excel/basicdatavalidation.md)|_Property_ > formula2|Gets or sets the Formula1, i.e. maximum value or value depending of the operator.|1.7D|
|[basicDataValidation](reference/excel/basicdatavalidation.md)|_Property_ > operator|The operator to use for validating the data.|1.7D|
|[customDataValidation](reference/excel/customdatavalidation.md)|_Property_ > formula|Custom data validation formula, it is to create special rules, such as preventing duplicates, or limiting the total in a range of cells.|1.7D|
|[customXmlPart](reference/excel/customxmlpart.md)|_Method_ > [deleteAttribute(xpath: string, namespaceMappings: object, name: string)]((reference/excel/customxmlpart.md#deleteattributexpath-string-namespacemappings-object-name-string)|Deletes an attribute with the given name from the element identified by xpath.|ApiSetAttribute.Spec|
|[customXmlPart](reference/excel/customxmlpart.md)|_Method_ > [deleteElement(xpath: string, namespaceMappings: object)]((reference/excel/customxmlpart.md#deleteelementxpath-string-namespacemappings-object)|Deletes the element identified by xpath.|ApiSetAttribute.Spec|
|[customXmlPart](reference/excel/customxmlpart.md)|_Method_ > [insertAttribute(xpath: string, namespaceMappings: object, name: string, value: string)]((reference/excel/customxmlpart.md#insertattributexpath-string-namespacemappings-object-name-string-value-string)|Inserts an attribute with the given name and value to the element identified by xpath.|ApiSetAttribute.Spec|
|[customXmlPart](reference/excel/customxmlpart.md)|_Method_ > [insertElement(xpath: string, xml: string, namespaceMappings: object, index: number)]((reference/excel/customxmlpart.md#insertelementxpath-string-xml-string-namespacemappings-object-index-number)|Inserts the given XML under the parent element identified by xpath at child position index.|ApiSetAttribute.Spec|
|[customXmlPart](reference/excel/customxmlpart.md)|_Method_ > [query(xpath: string, namespaceMappings: object)]((reference/excel/customxmlpart.md#queryxpath-string-namespacemappings-object)|Queries the XML content.|ApiSetAttribute.Spec|
|[customXmlPart](reference/excel/customxmlpart.md)|_Method_ > [updateAttribute(xpath: string, namespaceMappings: object, name: string, value: string)]((reference/excel/customxmlpart.md#updateattributexpath-string-namespacemappings-object-name-string-value-string)|Updates the value of an attribute with the given name of the element identified by xpath.|ApiSetAttribute.Spec|
|[customXmlPart](reference/excel/customxmlpart.md)|_Method_ > [updateElement(xpath: string, xml: string, namespaceMappings: object)]((reference/excel/customxmlpart.md#updateelementxpath-string-xml-string-namespacemappings-object)|Updates the XML of the element identified by xpath.|ApiSetAttribute.Spec|
|[dataValidation](reference/excel/datavalidation.md)|_Property_ > ignoreBlanks|Ignore blanks: no data validation will be performed on blank cells, it defaults to true.|1.7D|
|[dataValidation](reference/excel/datavalidation.md)|_Property_ > type|Type of the data validation, see Excel.DataValidationType for details. Read-only.|1.7D|
|[dataValidation](reference/excel/datavalidation.md)|_Relationship_ > errorAlert|Error alert when user enters invalid data.|1.7D|
|[dataValidation](reference/excel/datavalidation.md)|_Relationship_ > prompt|Prompt when users select a cell.|1.7D|
|[dataValidation](reference/excel/datavalidation.md)|_Relationship_ > rule|Data Validation rule that contains different type of data validation criteria.|1.7D|
|[dataValidation](reference/excel/datavalidation.md)|_Method_ > [clear()]((reference/excel/datavalidation.md#clear)|Clears the data validation from the current range.|1.7D|
|[dataValidationErrorAlert](reference/excel/datavalidationerroralert.md)|_Property_ > message|Represents error alert message.|1.7D|
|[dataValidationErrorAlert](reference/excel/datavalidationerroralert.md)|_Property_ > showAlert|It determines show error alert dialog or not when users enter invalid data, it defaults to true.|1.7D|
|[dataValidationErrorAlert](reference/excel/datavalidationerroralert.md)|_Property_ > style|Represents Data validation alert type, please see Excel.DataValidationAlertStyle for details.|1.7D|
|[dataValidationErrorAlert](reference/excel/datavalidationerroralert.md)|_Property_ > title|Represents error alert dialog title.|1.7D|
|[dataValidationPrompt](reference/excel/datavalidationprompt.md)|_Property_ > message|Represents the message of the prompt.|1.7D|
|[dataValidationPrompt](reference/excel/datavalidationprompt.md)|_Property_ > showPrompt|It determines showing the prompt or not when user selects a cell with the data validation.|1.7D|
|[dataValidationPrompt](reference/excel/datavalidationprompt.md)|_Property_ > title|Represents the title for the prompt.|1.7D|
|[dataValidationRule](reference/excel/datavalidationrule.md)|_Relationship_ > custom|Custom data validation criteria.|1.7D|
|[dataValidationRule](reference/excel/datavalidationrule.md)|_Relationship_ > date|Date data validation criteria.|1.7D|
|[dataValidationRule](reference/excel/datavalidationrule.md)|_Relationship_ > decimal|Decimal data validation criteria.|1.7D|
|[dataValidationRule](reference/excel/datavalidationrule.md)|_Relationship_ > list|List data validation criteria.|1.7D|
|[dataValidationRule](reference/excel/datavalidationrule.md)|_Relationship_ > textLength|TextLength data validation criteria.|1.7D|
|[dataValidationRule](reference/excel/datavalidationrule.md)|_Relationship_ > time|Time data validation criteria.|1.7D|
|[dataValidationRule](reference/excel/datavalidationrule.md)|_Relationship_ > wholeNumber|WholeNumber data validation criteria.|1.7D|
|[dateTimeDataValidation](reference/excel/datetimedatavalidation.md)|_Property_ > formula1|Gets or sets the Formula1, i.e. minimum value or value depending of the operator.|1.7D|
|[dateTimeDataValidation](reference/excel/datetimedatavalidation.md)|_Property_ > formula2|Gets or sets the Formula2, i.e. maximum value or value depending of the operator.|1.7D|
|[dateTimeDataValidation](reference/excel/datetimedatavalidation.md)|_Property_ > operator|The operator to use for validating the data.|1.7D|
|[listDataValidation](reference/excel/listdatavalidation.md)|_Property_ > inCellDropDown|Displays the list in cell drop down or not, it defaults to true.|1.7D|
|[listDataValidation](reference/excel/listdatavalidation.md)|_Property_ > source|Source of the list for data validation|1.7D|
|[namedItem](reference/excel/nameditem.md)|_Property_ > formula|Gets or sets the formula of the named item.  Formula always starts with a '=' sign.|1.7A|
|[namedItem](reference/excel/nameditem.md)|_Relationship_ > arrayValues|Returns an object containing values and types of the named item. Read-only.|1.7A|
|[namedItemArrayValues](reference/excel/nameditemarrayvalues.md)|_Property_ > types|Represents the types for each item in the named item array Read-only. Possible values are: Unknown, Empty, String, Integer, Double, Boolean, Error.|1.7A|
|[namedItemArrayValues](reference/excel/nameditemarrayvalues.md)|_Property_ > values|Represents the values of each item in the named item array. Read-only.|1.7A|
|[range](reference/excel/range.md)|_Property_ > isEntireColumn|Represents if the current range is an entire column. Read-only.|SM|
|[range](reference/excel/range.md)|_Property_ > isEntireRow|Represents if the current range is an entire row. Read-only.|SM|
|[range](reference/excel/range.md)|_Property_ > style|Represents the style of the current range. If the styles of the cells are inconsistent, null will be returned.|ApiSet.InProgressFeatures.Style|
|[range](reference/excel/range.md)|_Relationship_ > dataValidation|Returns a data validation object. Read-only.|1.7D|
|[range](reference/excel/range.md)|_Relationship_ > hyperlink|Represents the hyperlink set for the current range.|SM|
|[range](reference/excel/range.md)|_Method_ > [getAbsoluteResizedRange(numRows: number, numColumns: number)]((reference/excel/range.md#getabsoluteresizedrangenumrows-number-numcolumns-number)|Gets a Range object with the same top-left cell as the current Range object, but with the specified numbers of rows and columns.|SM|
|[range](reference/excel/range.md)|_Method_ > [getColumnsAfter(count: number)]((reference/excel/range.md#getcolumnsaftercount-number)|Gets a certain number of columns to the right of the current Range object.|ApiSet.PolyfillableDownTo1_1|
|[range](reference/excel/range.md)|_Method_ > [getColumnsBefore(count: number)]((reference/excel/range.md#getcolumnsbeforecount-number)|Gets a certain number of columns to the left of the current Range object.|ApiSet.PolyfillableDownTo1_1|
|[range](reference/excel/range.md)|_Method_ > [getResizedRange(deltaRows: number, deltaColumns: number)]((reference/excel/range.md#getresizedrangedeltarows-number-deltacolumns-number)|Gets a Range object similar to the current Range object, but with its bottom-right corner expanded (or contracted) by some number of rows and columns.|ApiSet.PolyfillableDownTo1_1|
|[range](reference/excel/range.md)|_Method_ > [getRowsAbove(count: number)]((reference/excel/range.md#getrowsabovecount-number)|Gets a certain number of rows above the current Range object.|ApiSet.PolyfillableDownTo1_1|
|[range](reference/excel/range.md)|_Method_ > [getRowsBelow(count: number)]((reference/excel/range.md#getrowsbelowcount-number)|Gets a certain number of rows below the current Range object.|ApiSet.PolyfillableDownTo1_1|
|[range](reference/excel/range.md)|_Method_ > [getSurroundingRegion()]((reference/excel/range.md#getsurroundingregion)|Returns a Range object that represents the surrounding region for the top-left cell in this range. A surrounding region is a range bounded by any combination of blank rows and blank columns relative to this range.|1.7|
|[range](reference/excel/range.md)|_Method_ > [showCard()]((reference/excel/range.md#showcard)|Displays the card for an active cell if it has rich value content.|1.7M|
|[rangeFormat](reference/excel/rangeformat.md)|_Property_ > textOrientation|Gets or sets the text orientation of all the cells within the range.|SM|
|[rangeHyperlink](reference/excel/rangehyperlink.md)|_Property_ > address|Represents the url target for the hyperlink.|SM|
|[rangeHyperlink](reference/excel/rangehyperlink.md)|_Property_ > documentReference|Represents the document reference target for the hyperlink.|SM|
|[rangeHyperlink](reference/excel/rangehyperlink.md)|_Property_ > screenTip|Represents the string displayed when hovering over the hyperlink.|SM|
|[rangeHyperlink](reference/excel/rangehyperlink.md)|_Property_ > textToDisplay|Represents the string that is displayed in the top left most cell in the range.|SM|
|[rangeViewCollection](reference/excel/rangeviewcollection.md)|_Method_ > [getFirst()]((reference/excel/rangeviewcollection.md#getfirst)|Gets the first RangeView object in the collection.|1.7GG|
|[rangeViewCollection](reference/excel/rangeviewcollection.md)|_Method_ > [getLast()]((reference/excel/rangeviewcollection.md#getlast)|Gets the last RangeView object in the collection.|1.7GG|
|[selectionChangedEventArgs](reference/excel/selectionchangedeventargs.md)|_Relationship_ > workbook|Gets the workbook object that raised the SelectionChanged event.|ApiSet.PolyfillableDownTo1_1|
|[tableColumn](reference/excel/tablecolumn.md)|_Method_ > [getNextColumn()]((reference/excel/tablecolumn.md#getnextcolumn)|Gets the table column that follows this one. If there are no table columns following this one, this method will throw an error.|ApiSet.InProgressFeatures.GetPreviousGetNext|
|[tableColumn](reference/excel/tablecolumn.md)|_Method_ > [getNextColumnOrNullObject()]((reference/excel/tablecolumn.md#getnextcolumnornullobject)|Gets the table column that follows this one. If there are no table columns following this one, this method will return a null object.|ApiSet.InProgressFeatures.GetPreviousGetNext|
|[tableColumn](reference/excel/tablecolumn.md)|_Method_ > [getPreviousColumn()]((reference/excel/tablecolumn.md#getpreviouscolumn)|Gets the table column that precedes this one. If there are no previous table columns, this method will throw an error.|ApiSet.InProgressFeatures.GetPreviousGetNext|
|[tableColumn](reference/excel/tablecolumn.md)|_Method_ > [getPreviousColumnOrNullObject()]((reference/excel/tablecolumn.md#getpreviouscolumnornullobject)|Gets the table column that precedes this one. If there are no previous table columns, this method will return a null objet.|ApiSet.InProgressFeatures.GetPreviousGetNext|
|[tableColumnCollection](reference/excel/tablecolumncollection.md)|_Method_ > [getFirst()]((reference/excel/tablecolumncollection.md#getfirst)|Gets the first column in the table.|1.7GG|
|[tableColumnCollection](reference/excel/tablecolumncollection.md)|_Method_ > [getLast()]((reference/excel/tablecolumncollection.md#getlast)|Gets the last column in the table.|1.7GG|
|[tableDataChangedEvent](reference/excel/tabledatachangedevent.md)|_Property_ > context|Gets the context object that facilitates requests to the Excel application.|1.7E|
|[tableDataChangedEvent](reference/excel/tabledatachangedevent.md)|_Relationship_ > changeType|Gets the change type that represents how the DataChanged event is triggered.|1.7E|
|[tableDataChangedEvent](reference/excel/tabledatachangedevent.md)|_Relationship_ > range|Gets the range that represents the changed area of a table on a specific worksheet.|1.7E|
|[tableDataChangedEvent](reference/excel/tabledatachangedevent.md)|_Relationship_ > source|Gets the source of the event.|1.7E|
|[tableDataChangedEvent](reference/excel/tabledatachangedevent.md)|_Relationship_ > table|Gets the table in which the data changed.|1.7E|
|[tableDataChangedEvent](reference/excel/tabledatachangedevent.md)|_Relationship_ > type|Gets the type of the event.|1.7E|
|[tableSelectionChangedEvent](reference/excel/tableselectionchangedevent.md)|_Property_ > context|Gets the context object that facilitates requests to the Excel application.|1.7E|
|[tableSelectionChangedEvent](reference/excel/tableselectionchangedevent.md)|_Relationship_ > range|Gets the range that represents the selected area of the table on a specific worksheet.|1.7E|
|[tableSelectionChangedEvent](reference/excel/tableselectionchangedevent.md)|_Relationship_ > table|Gets the table in which the selection changed.|1.7E|
|[tableSelectionChangedEvent](reference/excel/tableselectionchangedevent.md)|_Relationship_ > type|Gets the type of the event.|1.7E|
|[workbook](reference/excel/workbook.md)|_Property_ > name|Gets the workbook name. Read-only.|SM|
|[workbook](reference/excel/workbook.md)|_Relationship_ > properties|Gets the workbook properties. Read-only.|1.7DPGA|
|[workbook](reference/excel/workbook.md)|_Relationship_ > styles|Represents a collection of styles associated with the workbook. Read-only.|ApiSet.InProgressFeatures.Style|
|[workbook](reference/excel/workbook.md)|_Method_ > [getActiveCell()]((reference/excel/workbook.md#getactivecell)|Gets the currently active cell from the workbook.|SM|
|[workbook](reference/excel/workbook.md)|_Method_ > [getRange(address: string)]((reference/excel/workbook.md#getrangeaddress-string)|Gets the range object specified by the address or name.|1.7WR|
|[worksheet](reference/excel/worksheet.md)|_Property_ > gridlines|Gets or sets the worksheet's gridlines flag.|SM|
|[worksheet](reference/excel/worksheet.md)|_Property_ > headings|Gets or sets the worksheet's headings flag.|SM|
|[worksheet](reference/excel/worksheet.md)|_Property_ > tabColor|Gets or sets the worksheet tab color.|SM|
|[worksheet](reference/excel/worksheet.md)|_Relationship_ > freezePanes|Gets an object that can be used to manipulate frozen panes on the worksheet Read-only.|SM|
|[worksheet](reference/excel/worksheet.md)|_Relationship_ > pageLayout|Gets the PageLayout object of the worksheet. Read-only.|SM|
|[worksheet](reference/excel/worksheet.md)|_Method_ > [getRangeByIndexes(startRow: number, startColumn: number, rowCount: number, columnCount: number)]((reference/excel/worksheet.md#getrangebyindexesstartrow-number-startcolumn-number-rowcount-number-columncount-number)|Gets the range object beginning at a particular row index and column index, and spanning a certain number of rows and columns.|SM|
|[worksheetActivatedEvent](reference/excel/worksheetactivatedevent.md)|_Property_ > context|Gets the context object that facilitates requests to the Excel application.|1.7E|
|[worksheetActivatedEvent](reference/excel/worksheetactivatedevent.md)|_Relationship_ > type|Gets the type of the event.|1.7E|
|[worksheetActivatedEvent](reference/excel/worksheetactivatedevent.md)|_Relationship_ > worksheet|Gets the worksheet that is activated.|1.7E|
|[worksheetAddedEvent](reference/excel/worksheetaddedevent.md)|_Property_ > context|Gets the context object that facilitates requests to the Excel application.|1.7E|
|[worksheetAddedEvent](reference/excel/worksheetaddedevent.md)|_Relationship_ > source|Gets the source of the event.|1.7E|
|[worksheetAddedEvent](reference/excel/worksheetaddedevent.md)|_Relationship_ > type|Gets the type of the event.|1.7E|
|[worksheetAddedEvent](reference/excel/worksheetaddedevent.md)|_Relationship_ > worksheet|Gets the worksheet that is added to the workbook.|1.7E|
|[worksheetDataChangedEvent](reference/excel/worksheetdatachangedevent.md)|_Property_ > context|Gets the context object that facilitates requests to the Excel application.|1.7E|
|[worksheetDataChangedEvent](reference/excel/worksheetdatachangedevent.md)|_Relationship_ > changeType|Gets the change type that represents how the DataChanged event is triggered.|1.7E|
|[worksheetDataChangedEvent](reference/excel/worksheetdatachangedevent.md)|_Relationship_ > range|Gets the range that represents the changed area of a specific worksheet.|1.7E|
|[worksheetDataChangedEvent](reference/excel/worksheetdatachangedevent.md)|_Relationship_ > source|Gets the source of the event.|1.7E|
|[worksheetDataChangedEvent](reference/excel/worksheetdatachangedevent.md)|_Relationship_ > type|Gets the type of the event.|1.7E|
|[worksheetDeactivatedEvent](reference/excel/worksheetdeactivatedevent.md)|_Property_ > context|Gets the context object that facilitates requests to the Excel application.|1.7E|
|[worksheetDeactivatedEvent](reference/excel/worksheetdeactivatedevent.md)|_Relationship_ > type|Gets the type of the event.|1.7E|
|[worksheetDeactivatedEvent](reference/excel/worksheetdeactivatedevent.md)|_Relationship_ > worksheet|Gets the worksheet that is deactivated.|1.7E|
|[worksheetFreezePanes](reference/excel/worksheetfreezepanes.md)|_Method_ > [freezeAt(frozenRange: object)]((reference/excel/worksheetfreezepanes.md#freezeatfrozenrange-object)|Sets the frozen cells in the active worksheet view.|SM|
|[worksheetFreezePanes](reference/excel/worksheetfreezepanes.md)|_Method_ > [freezeColumns(count: number)]((reference/excel/worksheetfreezepanes.md#freezecolumnscount-number)|Freeze the first column(s) of the worksheet in place.|SM|
|[worksheetFreezePanes](reference/excel/worksheetfreezepanes.md)|_Method_ > [freezeRows(count: number)]((reference/excel/worksheetfreezepanes.md#freezerowscount-number)|Freeze the top row(s) of the worksheet in place.|SM|
|[worksheetFreezePanes](reference/excel/worksheetfreezepanes.md)|_Method_ > [getLocation()]((reference/excel/worksheetfreezepanes.md#getlocation)|Gets a range that describes the frozen cells in the active worksheet view.|SM|
|[worksheetFreezePanes](reference/excel/worksheetfreezepanes.md)|_Method_ > [getLocationOrNullObject()]((reference/excel/worksheetfreezepanes.md#getlocationornullobject)|Gets a range that describes the frozen cells in the active worksheet view.|SM|
|[worksheetFreezePanes](reference/excel/worksheetfreezepanes.md)|_Method_ > [unfreeze()]((reference/excel/worksheetfreezepanes.md#unfreeze)|Removes all frozen panes in the worksheet.|SM|
|[worksheetProtectionOptions](reference/excel/worksheetprotectionoptions.md)|_Property_ > allowEditObjects|Represents the worksheet protection option of allowing editing objects.|1.7W|
|[worksheetProtectionOptions](reference/excel/worksheetprotectionoptions.md)|_Property_ > allowEditScenarios|Represents the worksheet protection option of allowing editing scenarios.|1.7W|
|[worksheetProtectionOptions](reference/excel/worksheetprotectionoptions.md)|_Property_ > selectionMode|Represents the worksheet protection option of selection mode.|1.7W|


## Chart API additions

_Details of the specific APIs are listed below. Please let us know what you think!_

|Object| What is new| Description |Requirement set|
|:----|:----|:----|:----|
|[chartPointsCollection](reference/excel/chartpointscollection.md)|_Method_ > [getFirst()](reference/excel/chartpointscollection.md#getfirst)|Gets the first point in the series.|1.7|
|[chartPointsCollection](reference/excel/chartpointscollection.md)|_Method_ > [getLast()](reference/excel/chartpointscollection.md#getlast)|Gets the last point in the series.|1.7|
|[chartSeriesCollection](reference/excel/chartseriescollection.md)|_Method_ > [getFirst()](reference/excel/chartseriescollection.md#getfirst)|Gets the first series in the collection.|1.7|
|[chartSeriesCollection](reference/excel/chartseriescollection.md)|_Method_ > [getLast()](reference/excel/chartseriescollection.md#getlast)|Gets the last series in the collection.|1.7|
|[chartaxis](reference/excel/chartaxis.md)|_Property_ > baseTimeUnit|Returns or sets the base unit for the specified category axis.|1.9|
|[chartaxis](reference/excel/chartaxis.md)|_Property_ > categoryType|Represents the category type of chart Axis.|1.9|
|[chartaxis](reference/excel/chartaxis.md)|_Property_ > customDisplayUnit|Represents the custom value of axis display unit.|1.9|
|[chartaxis](reference/excel/chartaxis.md)|_Property_ > displayUnit|Represents the axis display unit of chart Axis. Read-only.|1.9|
|[chartaxis](reference/excel/chartaxis.md)|_Property_ > majorTimeUnitScale|Returns or sets the major unit scale value for the category axis when the CategoryType property is set to TimeScale.|1.9|
|[chartaxis](reference/excel/chartaxis.md)|_Property_ > minorTimeUnitScale|Returns or sets the minor unit scale value for the category axis when the CategoryType property is set to TimeScale.|1.9|
|[chartaxis](reference/excel/chartaxis.md)|_Property_ > showDisplayUnitLabel|Represents if show displayUnitLabel.|1.9|
|[chartaxis](reference/excel/chartaxis.md)|_Property_ > type|Represents the axis type. Read-only.|1.9|
|[chartaxis](reference/excel/chartaxis.md)|_Method_> [ setCategoryNames(sourceData:Range or string[])](reference/excel/chartaxis.md#setcategorynamessourcedata-rangeorstring[])|Sets all the category names for the specified axis.|1.9|
|[chartborder](reference/excel/chartborder.md)|_Method_ > [clear()](reference/excel/chartborder.md#clear)|Clear the border color of a chart element.|1.9|
|[chartborder](reference/excel/chartborder.md)|_Method_ > [setSolidColor(color: string)](reference/excel/chartborder.md#setsolidcolorcolor-string)|Sets the border formatting of a chart element to a uniform color.|1.9|
|[chartformatstring](reference/excel/chartformatstring.md)|_Relationship_ > [ChartFont](reference/excel/chartfont.md)|Represents the font properties of chartFormatString object. Read-only.|1.9|
|[chartpointformat](reference/excel/chartpointformat.md)|_Relationship_ > [ChartBorder](reference/excel/chartborder.md)|Represents the border format of a chart point, which includes border formating information. Read-only.|1.9|
|[chartseries](reference/excel/chartseries.md)|_Method_ > [delete()](reference/excel/chartseries.md#delete)|Deletes the chart series.|1.9|
|[chartseries](reference/excel/chartseries.md)|_Method_ > [setBubbleSizes(sourceData: Range)](reference/excel/chartseries.md#setbubblesizessourcedata-range)|Set bubble sizes for a chart series. Only works for bubble charts.|1.9|
|[chartseries](reference/excel/chartseries.md)|_Method_ > [setValues(sourceData: Range)](reference/excel/chartseries.md#setvaluessourcedata-range)|Set values for a chart series. For scatter chart, it means Y axis values.|1.9|
|[chartseries](reference/excel/chartseries.md)|_Method_ > [setXAxisValues(sourceData: Range)](reference/excel/chartseries.md#setxaxisvaluessourcedata-range)|Set values of X axis for a chart series. Only works for scatter charts.|1.9|
|[chartseries](reference/excel/chartseries.md)|_Relationship_ > [ChartTrendlineCollection](reference/excel/charttrendlinecollection.md)|Represents a collection of Trendlines in the series. Read-only.|1.9|
|[chartseriescollection](reference/excel/chartseriescollection.md)|_Method_ > [add(name: string, index: number)](reference/excel/chartseriescollection.md#addname-string-index-number)|Add a new series to the collection.|1.9|
|[chartseriesformat](reference/excel/chartseriesformat.md)|_Relationship_ > [ChartBorder](reference/excel/chartborder.md)|Represents the border format of a chart series, which includes border formating information. Read-only.|1.9|
|[charttitle](reference/excel/charttitle.md)|_Method_ > [getSubstring(start: number, length: number)](reference/excel/charttitle.md#getsubstringstart-number-length-number)|Get the characters of a chart title. Line break '\n' also counts one charater.|1.9|
|[charttitle](reference/excel/charttitle.md)|_Property_ > horizontalAlignment|Represents the horizontal alignment for chart title.|1.9|
|[charttrendline](reference/excel/charttrendline.md)|_Property_ > movingAverage|Represents the period of a chart trendline, only for trendline with MovingAverage type.|1.9|
|[charttrendline](reference/excel/charttrendline.md)|_Property_ > polynomialOrder|Represents the order of a chart trendline, only for trendline with Polynomial type.|1.9|
|[charttrendline](reference/excel/charttrendline.md)|_Property_ > type|Represents the type of a chart trendline. Possible values are: Linear, Expontential, Logarithmic, MovingAverage, Polynomial, Power.|1.9|
|[charttrendline](reference/excel/charttrendline.md)|_Relationship_ > [ChartTrendlineFormat](reference/excel/charttrendline.md)|Represents the formatting of a chart trendline. Read-only.|1.9|
|[charttrendlinecollection](reference/excel/charttrendlinecollection.md)|_Method_ > [add(type: string)](reference/excel/charttrendlinecollection.md#addtype-string)|Adds a new trendline to trendline collection.|1.9|
|[charttrendlinecollection](reference/excel/charttrendlinecollection.md)|_Method_ > [getCount](reference/excel/charttrendlinecollection.md#getcount)|Returns the number of trendlines in the collection.|1.9|
|[charttrendlinecollection](reference/excel/charttrendlinecollection.md)|_Method_ > [getItem(index: int)](reference/excel/charttrendlinecollection.md#getitemindex-int)|Get trendline object by index, which is the insertion order in items array.|1.9|
|[charttrendlinecollection](reference/excel/charttrendlinecollection.md)|_Property_ > items|A collection of chartTrendline objects. Read-only.|1.9|
|[charttrendlineformat](reference/excel/charttrendlineformat.md)|_Relationship_ > [ChartLineFormat](reference/excel/chartlineformat.md)|Represents chart line formatting. Read-only.|1.9|


## Give feedback

We need it, you want to give it. Feedback is much easier to give now that we're on GitHub. Check out the docs and let us know your feedback by submitting [issues](https://github.com/OfficeDev/office-js-docs/issues) directly in this repository.

For API support, you can post questions to the community on [StackOverflow](http://stackoverflow.com/questions/tagged/office-js) and tag them with [office-js].
