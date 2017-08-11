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

## Upcoming Excel 1.9 Release Features
* ChartTrendline
* Set Format to Chart Point(s), Subtitle 
* Catagory Axis
* Display Unit 
* ChartSeries Add/Delete/Modification
* Data validation (API to be updated..)

## Upcoming Excel Pivot1.1 Release Features

* Pivot table add/modify APIs (See the [bottom table](#upcoming-pivot-table-apis))


_Details of the specific APIs are listed below. Please let us know what you think!_

|Object| What is new| Description |Requirement set|
|:----|:----|:----|:----|
|[chartPointsCollection](reference/excel/chartpointscollection.md)|_Method_ > [getFirst()](reference/excel/chartpointscollection.md#getfirst)|Gets the first point in the series.|1.7|
|[chartPointsCollection](reference/excel/chartpointscollection.md)|_Method_ > [getLast()](reference/excel/chartpointscollection.md#getlast)|Gets the last point in the series.|1.7|
|[chartSeriesCollection](reference/excel/chartseriescollection.md)|_Method_ > [getFirst()](reference/excel/chartseriescollection.md#getfirst)|Gets the first series in the collection.|1.7|
|[chartSeriesCollection](reference/excel/chartseriescollection.md)|_Method_ > [getLast()](reference/excel/chartseriescollection.md#getlast)|Gets the last series in the collection.|1.7|
|[customXmlPart](reference/excel/customxmlpart.md)|_Method_ > [deleteAttribute(xpath: string, namespaceMappings: object, name: string)](reference/excel/customxmlpart.md#deleteattributexpath-string-namespacemappings-object-name-string)|Deletes an attribute with the given name from the element identified by xpath.|1.8|
|[customXmlPart](reference/excel/customxmlpart.md)|_Method_ > [deleteElement(xpath: string, namespaceMappings: object)](reference/excel/customxmlpart.md#deleteelementxpath-string-namespacemappings-object)|Deletes the element identified by xpath.|1.8|
|[customXmlPart](reference/excel/customxmlpart.md)|_Method_ > [insertAttribute(xpath: string, namespaceMappings: object, name: string, value: string)](reference/excel/customxmlpart.md#insertattributexpath-string-namespacemappings-object-name-string-value-string)|Inserts an attribute with the given name and value to the element identified by xpath.|1.8|
|[customXmlPart](reference/excel/customxmlpart.md)|_Method_ > [insertElement(xpath: string, xml: string, namespaceMappings: object, index: number)](reference/excel/customxmlpart.md#insertelementxpath-string-xml-string-namespacemappings-object-index-number)|Inserts the given XML under the parent element identified by xpath at child position index.|1.8|
|[customXmlPart](reference/excel/customxmlpart.md)|_Method_ > [query(xpath: string, namespaceMappings: object)](reference/excel/customxmlpart.md#queryxpath-string-namespacemappings-object)|Queries the XML content.|1.8|
|[customXmlPart](reference/excel/customxmlpart.md)|_Method_ > [updateAttribute(xpath: string, namespaceMappings: object, name: string, value: string)](reference/excel/customxmlpart.md#updateattributexpath-string-namespacemappings-object-name-string-value-string)|Updates the value of an attribute with the given name of the element identified by xpath.|1.8|
|[customXmlPart](reference/excel/customxmlpart.md)|_Method_ > [updateElement(xpath: string, xml: string, namespaceMappings: object)](reference/excel/customxmlpart.md#updateelementxpath-string-xml-string-namespacemappings-object)|Updates the XML of the element identified by xpath.|1.8|
|[range](reference/excel/range.md)|_Property_ > isEntireColumn|Represents if the current range is an entire column. Read-only.|1.7|
|[range](reference/excel/range.md)|_Property_ > isEntireRow|Represents if the current range is an entire row. Read-only.|1.7|
|[range](reference/excel/range.md)|_Relationship_ > hyperlink|Represents the hyperlink set for the current range.|1.7|
|[range](reference/excel/range.md)|_Method_ > [getAbsoluteResizedRange(numRows: number, numColumns: number)](reference/excel/range.md#getabsoluteresizedrangenumrows-number-numcolumns-number)|Gets a Range object with the same top-left cell as the current Range object, but with the specified numbers of rows and columns.|1.7|
|[range](reference/excel/range.md)|_Method_ > [getSurroundingRegion()](reference/excel/range.md#getsurroundingregion)|Returns a Range object that represents the surrounding region for this range. A surrounding region is a range bounded by any combination of blank rows and blank columns relative to this range.|1.7|
|[rangeFormat](reference/excel/rangeformat.md)|_Property_ > textOrientation|Gets or sets the text orientation of all the cells within the range.|1.7|
|[rangeHyperlink](reference/excel/rangehyperlink.md)|_Property_ > address|Represents the url target for the hyperlink.|1.7|
|[rangeHyperlink](reference/excel/rangehyperlink.md)|_Property_ > documentReference|Represents the document reference target for the hyperlink.|1.7|
|[rangeHyperlink](reference/excel/rangehyperlink.md)|_Property_ > screenTip|Represents the string displayed when hovering over the hyperlink.|1.7|
|[rangeHyperlink](reference/excel/rangehyperlink.md)|_Property_ > textToDisplay|Represents the string that is displayed in the top left most cell in the range.|1.7|
|[rangeViewCollection](reference/excel/rangeviewcollection.md)|_Method_ > [getFirst()](reference/excel/rangeviewcollection.md#getfirst)|Gets the first RangeView object in the collection.|1.7|
|[rangeViewCollection](reference/excel/rangeviewcollection.md)|_Method_ > [getLast()](reference/excel/rangeviewcollection.md#getlast)|Gets the last RangeView object in the collection.|1.7|
|[tableColumn](reference/excel/tablecolumn.md)|_Method_ > [getNextColumn()](reference/excel/tablecolumn.md#getnextcolumn)|Gets the table column that follows this one. If there are no table columns following this one, this method will throw an error.|1.7|
|[tableColumn](reference/excel/tablecolumn.md)|_Method_ > [getNextColumnOrNullObject()](reference/excel/tablecolumn.md#getnextcolumnornullobject)|Gets the table column that follows this one. If there are no table columns following this one, this method will return a null object.|1.7|
|[tableColumn](reference/excel/tablecolumn.md)|_Method_ > [getPreviousColumn()](reference/excel/tablecolumn.md#getpreviouscolumn)|Gets the table column that precedes this one. If there are no previous table columns, this method will throw an error.|1.7|
|[tableColumn](reference/excel/tablecolumn.md)|_Method_ > [getPreviousColumnOrNullObject()](reference/excel/tablecolumn.md#getpreviouscolumnornullobject)|Gets the table column that precedes this one. If there are no previous table columns, this method will return a null objet.|1.7|
|[tableColumnCollection](reference/excel/tablecolumncollection.md)|_Method_ > [getFirst()](reference/excel/tablecolumncollection.md#getfirst)|Gets the first column in the table.|1.7|
|[tableColumnCollection](reference/excel/tablecolumncollection.md)|_Method_ > [getLast()](reference/excel/tablecolumncollection.md#getlast)|Gets the last column in the table.|1.7|
|[workbook](reference/excel/workbook.md)|_Property_ > name|Gets the workbook name. Read-only.|1.7|
|[workbook](reference/excel/workbook.md)|_Method_ > [getRange(address: string)](reference/excel/workbook.md#getrangeaddress-string)|Gets the range object specified by the address or name.|1.7|
|[worksheet](reference/excel/worksheet.md)|_Property_ > gridlines|Gets or sets the worksheet's gridlines flag.|1.7|
|[worksheet](reference/excel/worksheet.md)|_Property_ > headings|Gets or sets the worksheet's headings flag.|1.7|
|[worksheet](reference/excel/worksheet.md)|_Property_ > tabColor|Gets or sets the worksheet tab color.|1.7|
|[worksheet](reference/excel/worksheet.md)|_Method_ > [getRangeByIndexes(startRow: number, startColumn: number, rowCount: number, columnCount: number)](reference/excel/worksheet.md#getrangebyindexesstartrow-number-startcolumn-number-rowcount-number-columncount-number)|Gets the range object beginning at a particular row index and column index, and spanning a certain number of rows and columns.|1.7|
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

## Upcoming Pivot Table APIs


|Object| What is new| Description |Requirement set|
|:----|:----|:----|:----|
|[calculatedFieldCollection](reference/excel/calculatedfieldcollection.md)|_Property_ > items|A collection of calculatedField objects. Read-only.|Pivot1.1|
|[calculatedFieldCollection](reference/excel/calculatedfieldcollection.md)|_Method_ > [add(Name: string, Formula: string, UseStandardFormula: bool)](reference/excel/calculatedfieldcollection.md#addname-string-formula-string-usestandardformula-bool)|Creates a new calculated field. Returns a PivotField object.|Pivot1.1|
|[calculatedFieldCollection](reference/excel/calculatedfieldcollection.md)|_Method_ > [getCount()](reference/excel/calculatedfieldcollection.md#getcount)|Returns the number of objects in the collection.|Pivot1.1|
|[calculatedFieldCollection](reference/excel/calculatedfieldcollection.md)|_Method_ > [getItem(nameOrIndex: object)](reference/excel/calculatedfieldcollection.md#getitemnameorindex-object)|Returns a single object from a collection based on name or index.|Pivot1.1|
|[calculatedFieldCollection](reference/excel/calculatedfieldcollection.md)|_Method_ > [getItemAt(index: number)](reference/excel/calculatedfieldcollection.md#getitematindex-number)|Returns a single object from a collection based on index.|Pivot1.1|
|[pivotApiInvokeHelper](reference/excel/pivotapiinvokehelper.md)|_Method_ > [pivotCacheCollectionAdd(pivotCaches: PivotCacheCollection, SourceType: string, address: Range)](reference/excel/pivotapiinvokehelper.md#pivotcachecollectionaddpivotcaches-pivotcachecollection-sourcetype-string-address-range)|Add a pivot cache|Pivot1.1|
|[pivotApiInvokeHelper](reference/excel/pivotapiinvokehelper.md)|_Method_ > [pivotFieldGetSubtotals(pivotField: PivotField)](reference/excel/pivotapiinvokehelper.md#pivotfieldgetsubtotalspivotfield-pivotfield)|Get pivot field sub totals|Pivot1.1|
|[pivotApiInvokeHelper](reference/excel/pivotapiinvokehelper.md)|_Method_ > [pivotFieldSetSubtotals(pivotField: PivotField, subtotals: Subtotals)](reference/excel/pivotapiinvokehelper.md#pivotfieldsetsubtotalspivotfield-pivotfield-subtotals-subtotals)|Set pivot field sub totals|Pivot1.1|
|[pivotApiInvokeHelper](reference/excel/pivotapiinvokehelper.md)|_Method_ > [pivotTableAddChart(pivotTable: PivotTableExt, chartType: string, seriesBy: string)](reference/excel/pivotapiinvokehelper.md#pivottableaddchartpivottable-pivottableext-charttype-string-seriesby-string)|Add a pivot chart|Pivot1.1|
|[pivotApiInvokeHelper](reference/excel/pivotapiinvokehelper.md)|_Method_ > [pivotTableCollectionAdd(pivotTables: PivotTableCollectionExt, name: string, address: Range, pivotCache: PivotCache)](reference/excel/pivotapiinvokehelper.md#pivottablecollectionaddpivottables-pivottablecollectionext-name-string-address-range-pivotcache-pivotcache)|Add a pivot table|Pivot1.1|
|[pivotCache](reference/excel/pivotcache.md)|_Property_ > id|Identifier for the specified PivotCache object. Read-only.|Pivot1.1|
|[pivotCache](reference/excel/pivotcache.md)|_Property_ > index|Returns a value that represents the index number of the object within the collection of similar objects. Read-only.|Pivot1.1|
|[pivotCache](reference/excel/pivotcache.md)|_Property_ > version|Returns the version of Microsoft Excel in which the PivotCache was created. Read-only.|Pivot1.1|
|[pivotCache](reference/excel/pivotcache.md)|_Method_ > [refresh()](reference/excel/pivotcache.md#refresh)|Causes the specified cache to be refreshed and associated chart to be redrawn immediately.|Pivot1.1|
|[pivotCacheCollection](reference/excel/pivotcachecollection.md)|_Property_ > items|A collection of pivotCache objects. Read-only.|Pivot1.1|
|[pivotCacheCollection](reference/excel/pivotcachecollection.md)|_Method_ > [add(sourceType: string, address: Range)](reference/excel/pivotcachecollection.md#addsourcetype-string-address-range)|Creates a new PivotCache.|Pivot1.1|
|[pivotCacheCollection](reference/excel/pivotcachecollection.md)|_Method_ > [getCount()](reference/excel/pivotcachecollection.md#getcount)|Returns a value that represents the number of objects in the collection.|Pivot1.1|
|[pivotCacheCollection](reference/excel/pivotcachecollection.md)|_Method_ > [getItem(index: number)](reference/excel/pivotcachecollection.md#getitemindex-number)|Returns a single object from a collection based on the index.|Pivot1.1|
|[pivotCacheCollection](reference/excel/pivotcachecollection.md)|_Method_ > [getItemAt(index: number)](reference/excel/pivotcachecollection.md#getitematindex-number)|Returns a single object from a collection based on the index.|Pivot1.1|
|[pivotField](reference/excel/pivotfield.md)|_Property_ > aggregationFunction|Returns or sets the function used to summarize the PivotTable field (data fields only).|Pivot1.1|
|[pivotField](reference/excel/pivotfield.md)|_Property_ > allItemsVisible|Used to retrieve a Boolean value that indicates whether or not any manual filtering is applied to the PivotField. Read-only.|Pivot1.1|
|[pivotField](reference/excel/pivotfield.md)|_Property_ > autoSortField|Returns the name of the data field used to sort the specified PivotTable field automatically. Read-only.|Pivot1.1|
|[pivotField](reference/excel/pivotfield.md)|_Property_ > autoSortOrder|Returns the order used to sort the specified PivotTable field automatically. Read-only.|Pivot1.1|
|[pivotField](reference/excel/pivotfield.md)|_Property_ > calculated|True if the PivotTable field is a calculated field or item. Read-only.|Pivot1.1|
|[pivotField](reference/excel/pivotfield.md)|_Property_ > calculation|Returns or sets a PivotFieldCalculation value that represents the type of calculation performed by the specified field. This property is valid only for data fields.|Pivot1.1|
|[pivotField](reference/excel/pivotfield.md)|_Property_ > caption|Returns a String value that represents the label text for the pivot field.|Pivot1.1|
|[pivotField](reference/excel/pivotfield.md)|_Property_ > dataType|Returns a PivotFieldDataType value that represents the type of data in the PivotTable field. Read-only.|Pivot1.1|
|[pivotField](reference/excel/pivotfield.md)|_Property_ > drilledDown|True if the flag for the specified PivotTable field or PivotTable item is set to drilled (expanded, or visible).|Pivot1.1|
|[pivotField](reference/excel/pivotfield.md)|_Property_ > enableMultiplePageItems|Used for specifying whether or not check boxes are present in the filter drop-down list for fields in the page area.|Pivot1.1|
|[pivotField](reference/excel/pivotfield.md)|_Property_ > formula|Returns or sets a string value that represents the object's formula in A1-style notation.|Pivot1.1|
|[pivotField](reference/excel/pivotfield.md)|_Property_ > name|Returns or sets a String value representing the name of the object.|Pivot1.1|
|[pivotField](reference/excel/pivotfield.md)|_Property_ > numberFormat|Returns or sets a String value that represents the format code for the object.|Pivot1.1|
|[pivotField](reference/excel/pivotfield.md)|_Property_ > orientation|Returns or sets a value that represents the location of the field in the specified PivotTable report.|Pivot1.1|
|[pivotField](reference/excel/pivotfield.md)|_Property_ > position|Returns or sets a value that represents the position of the field (first, second, third, and so on) among all the fields in its orientation (Rows, Columns, Pages, Data).|Pivot1.1|
|[pivotField](reference/excel/pivotfield.md)|_Property_ > showDetail|Gets or sets whether the specified PivotField is showing detail.|Pivot1.1|
|[pivotField](reference/excel/pivotfield.md)|_Property_ > sourceName|Returns a string value that represents the specified objectname as it appears in the original source data for the specified PivotTable report. Read-only.|Pivot1.1|
|[pivotField](reference/excel/pivotfield.md)|_Relationship_ > currentPage|Returns or sets the current page showing for the page field (valid only for page fields).|Pivot1.1|
|[pivotField](reference/excel/pivotfield.md)|_Relationship_ > hiddenItems|Returns an object that represents either a single hidden PivotTable item (a PivotItemobject) or a collection of all the hidden items (a PivotItems object) in the specified field. Read-only.|Pivot1.1|
|[pivotField](reference/excel/pivotfield.md)|_Relationship_ > pivotItems|Returns an object that represents either a single PivotTable item (a PivotItem object) or a collection of all the visible and hidden items (a PivotItems object) in the specified field. Read-only.|Pivot1.1|
|[pivotField](reference/excel/pivotfield.md)|_Relationship_ > subtotals|Returns or sets subtotals displayed with the specified field. Valid only for nondata fields.|Pivot1.1|
|[pivotField](reference/excel/pivotfield.md)|_Relationship_ > visiblePivotItems|Returns an object that represents either a single visible PivotTable item (a PivotItemobject) or a collection of all the visible items (a PivotItems object) in the specified field. Read-only.|Pivot1.1|
|[pivotField](reference/excel/pivotfield.md)|_Method_ > [autoGroup()](reference/excel/pivotfield.md#autogroup)|Automatically groups the pivot fields in a pivot table.|Pivot1.1|
|[pivotField](reference/excel/pivotfield.md)|_Method_ > [autoSort(sortOrder: string, Field: string)](reference/excel/pivotfield.md#autosortsortorder-string-field-string)|Establishes automatic field-sorting rules for PivotTable reports.|Pivot1.1|
|[pivotField](reference/excel/pivotfield.md)|_Method_ > [clearAllFilters()](reference/excel/pivotfield.md#clearallfilters)|Calling this method deletes all filters currently applied to the PivotField. This includes deleting all filters from the PivotFilters collection of the PivotField and removing any manual filtering applied to the PivotField as well. If the PivotField is in the Report Filter area, the item selected will be set to the default item.|Pivot1.1|
|[pivotField](reference/excel/pivotfield.md)|_Method_ > [getDataRange()](reference/excel/pivotfield.md#getdatarange)|Returns a Range object that represents the range of the current PivotField.|Pivot1.1|
|[pivotFieldCollection](reference/excel/pivotfieldcollection.md)|_Property_ > items|A collection of pivotField objects. Read-only.|Pivot1.1|
|[pivotFieldCollection](reference/excel/pivotfieldcollection.md)|_Method_ > [getCount()](reference/excel/pivotfieldcollection.md#getcount)|Returns a value that represents the number of objects in the collection.|Pivot1.1|
|[pivotFieldCollection](reference/excel/pivotfieldcollection.md)|_Method_ > [getItem(nameOrIndex: object)](reference/excel/pivotfieldcollection.md#getitemnameorindex-object)|Returns a single object from a collection based on name or index.|Pivot1.1|
|[pivotFieldCollection](reference/excel/pivotfieldcollection.md)|_Method_ > [getItemAt(index: number)](reference/excel/pivotfieldcollection.md#getitematindex-number)|Returns a single object from a collection based on the index.|Pivot1.1|
|[pivotItem](reference/excel/pivotitem.md)|_Property_ > calculated|True if the PivotTable item is a calculated field or item. Read-only.|Pivot1.1|
|[pivotItem](reference/excel/pivotitem.md)|_Property_ > drilledDown|True if the flag for the specified PivotTable field or PivotTable item is set to drilled (expanded, or visible).|Pivot1.1|
|[pivotItem](reference/excel/pivotitem.md)|_Property_ > name|Returns or sets a String value representing the name of the object.|Pivot1.1|
|[pivotItem](reference/excel/pivotitem.md)|_Property_ > position|Returns or sets a Long value that represents the position of the item in its field, if the item is currently showing.|Pivot1.1|
|[pivotItem](reference/excel/pivotitem.md)|_Property_ > recordCount|Returns the number of records in the PivotTable cache or the number of cache records that contain the specified item. Read-only.|Pivot1.1|
|[pivotItem](reference/excel/pivotitem.md)|_Property_ > showDetail|True if the outline is expanded for the specified range (so that the detail of the column or row is visible). The specified range must be a single summary column or row in an outline. For the PivotItem object (or the Range object if the range is in a PivotTable report), this property is set to True if the item is showing detail.|Pivot1.1|
|[pivotItem](reference/excel/pivotitem.md)|_Property_ > sourceName|Returns a value that represents the specified objectname, as it appears in the original source data, for the specified PivotTable report. Read-only.|Pivot1.1|
|[pivotItem](reference/excel/pivotitem.md)|_Property_ > visible|Returns or sets a Boolean value that determines whether the object is visible.|Pivot1.1|
|[pivotItem](reference/excel/pivotitem.md)|_Relationship_ > pivotField|Returns the parent Pivot Field for the current PivotItem. Read-only.|Pivot1.1|
|[pivotItem](reference/excel/pivotitem.md)|_Method_ > [getDataRange()](reference/excel/pivotitem.md#getdatarange)|Returns a Range object that represents the range of the current PivotItem.|Pivot1.1|
|[pivotItemCollection](reference/excel/pivotitemcollection.md)|_Property_ > items|A collection of pivotItem objects. Read-only.|Pivot1.1|
|[pivotItemCollection](reference/excel/pivotitemcollection.md)|_Method_ > [getCount()](reference/excel/pivotitemcollection.md#getcount)|Returns a value that represents the number of objects in the collection.|Pivot1.1|
|[pivotItemCollection](reference/excel/pivotitemcollection.md)|_Method_ > [getItem(nameOrIndex: object)](reference/excel/pivotitemcollection.md#getitemnameorindex-object)|Returns a single object from a collection based on name or index.|Pivot1.1|
|[pivotItemCollection](reference/excel/pivotitemcollection.md)|_Method_ > [getItemAt(index: number)](reference/excel/pivotitemcollection.md#getitematindex-number)|Returns a single object from a collection based on the index.|Pivot1.1|
|[pivotTableCollectionExt](reference/excel/pivottablecollectionext.md)|_Property_ > items|A collection of pivotTable objects. Read-only.|Pivot1.1|
|[pivotTableCollectionExt](reference/excel/pivottablecollectionext.md)|_Method_ > [add(name: string, address: Range, pivotCache: PivotCache)](reference/excel/pivottablecollectionext.md#addname-string-address-range-pivotcache-pivotcache)|Adds a new PivotTable report. Returns a PivotTable object.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Property_ > allowMultipleFilters|Sets or retrieves a value that indicates whether a PivotField can have multiple filters applied to it at the same time.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Property_ > alternativeText|Returns or sets the descriptive (alternative) text string for the specified PivotTable.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Property_ > columnGrandTotals|True if the PivotTable report shows grand totals for columns.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Property_ > compactLayoutColumnHeader|Specifies the caption that is displayed in the column header of a PivotTable when in compact row layout form.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Property_ > compactLayoutRowHeader|Specifies the caption that is displayed in the row header of a PivotTable when in compact row layout form.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Property_ > compactRowIndent|Returns or sets the indent increment for PivotItems when compact row layout form is turned on.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Property_ > displayContextTooltips|Controls whether or not tooltips are displayed for PivotTable cells.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Property_ > displayEmptyColumn|Returns True when the non-empty MDX keyword is included in the query to the OLAP provider for the value axis. The OLAP provider will not return empty columns in the result set. Returns False when the non-empty keyword is omitted.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Property_ > displayEmptyRow|Returns True when the non-empty MDX keyword is included in the query to the OLAP provider for the category axis. The OLAP provider will not return empty rows in the result set. Returns False when the non-empty keyword is omitted.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Property_ > displayErrorString|True if the PivotTable report displays a custom error string in cells that contain errors. The default value is False.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Property_ > displayFieldCaptions|Controls whether or not filter buttons and PivotField captions for rows and columns are displayed in the grid.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Property_ > displayNullString|True if the PivotTable report displays a custom string in cells that contain null values. The default value is True.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Property_ > enableDrilldown|True if drilldown is enabled. The default value is True.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Property_ > enableFieldDialog|True if the PivotTable Field dialog box is available when the user double-clicks the PivotTable field. The default value is True.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Property_ > enableFieldList|False to disable the ability to display the field list for the PivotTable. If the field list was already being displayed it disappears. The default value is True.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Property_ > enableWizard|True if the PivotTable Wizard is available. The default value is True.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Property_ > errorString|Returns or sets a String value that represents the string displayed in cells that contain errors when the DisplayErrorString property is True.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Property_ > fieldListSortAscending|Controls the sort order of fields in the PivotTable Field List. When this property is set toTrue, the fields are sorted in ascending order. When it is set to False, the fields are sorted in data source order.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Property_ > grandTotalName|Returns or sets the text string label that is displayed in the grand total column or row heading in the specified PivotTable report. The default value is the string `Grand Total`.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Property_ > hasAutoFormat|True if the PivotTable report is automatically formatted when itrefreshed or when fields are moved.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Property_ > hidden|Checks whether the PivotTable exists at the worksheet level. Boolean. Read-only.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Property_ > nullString|Returns or sets the string displayed in cells that contain null values when theDisplayNullString property is True. The default value is an empty string.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Property_ > preserveFormatting|True if formatting is preserved when the report is refreshed or recalculated by operations such as pivoting, sorting, or changing page field items.For query tables, this property is True if any formatting common to the first five rows of data are applied to new rows of data in the query table. Unused cells aren╬ô├ç├ût formatted. The property isFalse if the last AutoFormat applied to the query table is applied to new rows of data. The default value is True.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Property_ > printDrillIndicators|Specifies whether or not drill indicators are printed with the PivotTable.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Property_ > printTitles|True if the print titles for the worksheet are set based on the PivotTable report. False if the print titles for the worksheet are used. The default value is False.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Property_ > refreshName|Returns the name of the person who last refreshed the PivotTable report data. Read-only.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Property_ > repeatItemsOnEachPrintedPage|True if row, column, and item labels appear on the first row of each page when the specified PivotTable report is printed. False if labels are printed only on the first page. The default value is True.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Property_ > rowGrandTotals|True if the PivotTable report shows grand totals for rows.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Property_ > saveData|True if data for the PivotTable report is saved with the workbook. False if only the report definition is saved.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Property_ > showDrillIndicators|The ShowDrillIndicators property is used for toggling the display of drill indicators in the PivotTable.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Property_ > showPageMultipleItemLabel|When set to True (default), Multiple items will appear in the PivotTable cell on the worksheet whenever items are hidden and an aggregate of non-hidden items is shown in the PivotTable view.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Property_ > showTableStyleColumnHeaders|The ShowTableStyleColumnHeaders property is set to True if the coulmn headers should be displayed in the PivotTable.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Property_ > showTableStyleColumnStripes|The ShowTableStyleColumnStripes property displays banded columns in which even columns are formatted differently from odd columns. This makes PivotTables easier to read.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Property_ > showTableStyleLastColumn|Returns or sets if the last column is displayed for the specified PivotTable object.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Property_ > showTableStyleRowHeaders|The ShowTableStyleRowHeaders property is set to True if the row headers should be displayed in the PivotTable.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Property_ > showTableStyleRowStripes|The ShowTableStyleRowStripes property displays banded rows in which even rows are formatted differently from odd rows. This makes PivotTables easier to read.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Property_ > showValuesRow|Returns or sets whether the values row is displayed.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Property_ > smallGrid|True if Microsoft Excel uses a grid thattwo cells wide and two cells deep for a newly created PivotTable report. False if Excel uses a blank stencil outline.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Property_ > sortUsingCustomLists|The SortUsingCustomLists property controls whether custom lists are used for sorting items of fields, both initially when the PivotField is initialized and the PivotItems are ordered by their captions; and later when the user applies a sort.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Property_ > subtotalHiddenPageItems|True if hidden page field items in the PivotTable report are included in row and column subtotals, block totals, and grand totals. The default value is False.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Property_ > summary|Returns or sets the description associated with the alternative text string for the specified PivotTable.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Property_ > tag|Returns or sets a string saved with the PivotTable report.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Property_ > totalsAnnotation|True if an asterisk (*) is displayed next to each subtotal and grand total value in the specified PivotTable report if the report is based on an OLAP data source. The default value is True.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Property_ > vacatedStyle|Returns or sets the style applied to cells vacated when the PivotTable report is refreshed. The default value is a null string (no style is applied by default).|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Property_ > value|Returns or sets a String value that represents the name of the PivotTable report.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Property_ > version|Returns a value that represents the Microsoft Excel version number. Read-only.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Relationship_ > calculatedFields|Returns a CalculatedFields collection that represents all the calculated fields in the specified PivotTable report. Read-only.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Relationship_ > dataBodyRange|Returns a Range object that represents the range of values in a PivotTable. Read-only.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Relationship_ > dataLabelRange|Returns a Range object that represents the range that contains the labels for the data fields in the PivotTable report. Read-only.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Relationship_ > pivotFields|Returns an object that represents either a single PivotTable field (a PivotField object) or a collection of both the visible and hidden fields (a PivotFields object) in the PivotTable report. Read-only.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Relationship_ > refreshDate|Returns the date on which the PivotTable report was last refreshed. Read-only.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Method_ > [addChart(chartType: string, seriesBy: string)](reference/excel/pivottableext.md#addchartcharttype-string-seriesby-string)|Creates a new chart.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Method_ > [addDataField(field: PivotField, caption: string, func: string)](reference/excel/pivottableext.md#adddatafieldfield-pivotfield-caption-string-func-string)|Adds a data field to a PivotTable report. Returns a PivotField object that represents the new data field.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Method_ > [clearTable()](reference/excel/pivottableext.md#cleartable)|The ClearTable method is used for clearing a PivotTable. Clearing PivotTables includes removing all the fields and deleting all filtering and sorting applied to the PivotTables. This method resets the PivotTable to the state it had right after it was created, before any fields were added to it.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Method_ > [getColumnField(Index: number)](reference/excel/pivottableext.md#getcolumnfieldindex-number)|Returns an object that represents either a single PivotTable field (a PivotField object) or a collection of all the fields (a PivotFields object) that are currently shown as column fields.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Method_ > [getColumnRange()](reference/excel/pivottableext.md#getcolumnrange)|Returns a Range object that represents the range that contains the column area in the PivotTable report.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Method_ > [getDataField(Index: number)](reference/excel/pivottableext.md#getdatafieldindex-number)|Returns an object that represents either a single PivotTable field (a PivotField object) or a collection of all the fields (a PivotFields object) that are currently shown as data fields.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Method_ > [getEntireRange()](reference/excel/pivottableext.md#getentirerange)|Returns a Range object that represents the range containing the entire PivotTable report, including page fields.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Method_ > [getHiddenField(index: number)](reference/excel/pivottableext.md#gethiddenfieldindex-number)|Returns an object that represents either a single PivotTable field (a PivotField object) or a collection of all the fields (a PivotFields object) that are currently not shown as row, column, page, or data fields.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Method_ > [getRowField(index: number)](reference/excel/pivottableext.md#getrowfieldindex-number)|Returns an object that represents either a single field in a PivotTable report (a PivotField object) or a collection of all the fields (a PivotFields object) that are currently showing as row fields.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Method_ > [getRowRange()](reference/excel/pivottableext.md#getrowrange)|Returns a Range object that represents the range including the row area on the PivotTable report.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Method_ > [listFormulas()](reference/excel/pivottableext.md#listformulas)|Creates a list of calculated PivotTable items and fields on a separate worksheet.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Method_ > [refreshTable()](reference/excel/pivottableext.md#refreshtable)|Refreshes the PivotTable report from the source data. Returns True if itsuccessful.|Pivot1.1|
|[pivotTableExt](reference/excel/pivottableext.md)|_Method_ > [update()](reference/excel/pivottableext.md#update)|Updates the PivotTable report.|Pivot1.1|
|[subtotals](reference/excel/subtotals.md)|_Property_ > automatic|If Automatic is set to true, then all other values will be ignored when setting the Subtotals.|Pivot1.1|
|[subtotals](reference/excel/subtotals.md)|_Property_ > average|Average|Pivot1.1|
|[subtotals](reference/excel/subtotals.md)|_Property_ > count|Count|Pivot1.1|
|[subtotals](reference/excel/subtotals.md)|_Property_ > countNumbers|CountNumbers|Pivot1.1|
|[subtotals](reference/excel/subtotals.md)|_Property_ > max|Max|Pivot1.1|
|[subtotals](reference/excel/subtotals.md)|_Property_ > min|Min|Pivot1.1|
|[subtotals](reference/excel/subtotals.md)|_Property_ > product|Product|Pivot1.1|
|[subtotals](reference/excel/subtotals.md)|_Property_ > standardDeviation|StandardDeviation|Pivot1.1|
|[subtotals](reference/excel/subtotals.md)|_Property_ > standardDeviationP|StandardDeviationP|Pivot1.1|
|[subtotals](reference/excel/subtotals.md)|_Property_ > sum|Sum|Pivot1.1|
|[subtotals](reference/excel/subtotals.md)|_Property_ > variation|Variation|Pivot1.1|
|[subtotals](reference/excel/subtotals.md)|_Property_ > variationP|VariationP|Pivot1.1|
|[workbookExt](reference/excel/workbookext.md)|_Relationship_ > pivotCaches|Returns a PivotCaches collection that represents all the PivotTable caches in the specified workbook. Read-only.|Pivot1.1|



## Give feedback

We need it, you want to give it. Feedback is much easier to give now that we're on GitHub. Check out the docs and let us know your feedback by submitting [issues](https://github.com/OfficeDev/office-js-docs/issues) directly in this repository.

For API support, you can post questions to the community on [StackOverflow](http://stackoverflow.com/questions/tagged/office-js) and tag them with [office-js].
