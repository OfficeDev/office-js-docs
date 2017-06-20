## Excel JavaScript API Open Specification

_Applies to: Excel 2016, Excel Online, Excel for iOS, Excel for Mac_

Below links describes the new set of Excel JavaScript APIs that are being planned for the next releases (Requirement Set 1.7 and beyond). Please review and provide your feedback. One of the best ways of providing your input is by opening new issue in GitHub using the links available below.

_**Note**: below listed features are still under design and review phase and hence not yet available as part of the product. The final design is subject to change. Once the feature is made available, the final specification will be published as part of the master repository._

## Upcoming Excel 1.7+ Release Features

* [Data validation](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_OpenSpec/reference/excel/datavalidation.md).
* Workbook name
* Gridlines
* Tab color
* Text orientation
* Worksheet headings
* New range functions
* Freeze rows/columns.
* ChartTrendline
* Set Format to Chart Point(s), Subtitle 
* Catagory Axis
* Display Unit 
* ChartSeries Add/Delete/Modification

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
|[chartaxis](reference/excel/chartaxis.md)|_Method_> [ setCategoryNames(catagoryNames:string[])](reference/excel/chartborder.md#setcategorynames-string)|Sets all the category names for the specified axis.|1.9|
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
