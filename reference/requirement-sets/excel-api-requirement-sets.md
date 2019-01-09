# Excel JavaScript API requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Specify Office hosts and API requirements](../../docs/overview/specify-office-hosts-and-api-requirements.md).

Excel add-ins run across multiple versions of Office, including Office 2016 for Windows, Office for iPad, Office for Mac, and Office Online. The following table lists the Excel requirement sets, the Office host applications that support that requirement set, and the build versions or number for those applications.

> **Note**: Any API that is listed as **beta** is not ready for production usage. They are made available so that developers can try them out in test and development environments. They are not meant to be used against production/business critical documents. 

> For the requirement sets that are marked as *Beta*, use the specified (or later) version of the Office software and use the Beta library of the CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js. Entires not listed as *Beta* are generally available and you can continue to use Production CDN library: https://appsforoffice.microsoft.com/lib/1/hosted/office.js

|  Requirement set  |  Office 365 for Windows\*  |  Office 365 for iPad  |  Office 365 for Mac  | Office Online  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|
| Beta  | Please [visit our Excel JavaScript API open specification page](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec)! |
| ExcelApi1.8  | Version 1808 (Build 10730.20102) or later| 2.17 or later | 16.17 or later | Sept 2018 | Coming soon |
| ExcelApi1.7  | Version 1801 (Build 9001.2171) or later| 2.9 or later | 16.9 or later | April 2018 | Coming soon |
| ExcelApi1.6  | Version 1704 (Build 8201.2001) or later| 2.2 or later |15.36 or later| April 2017 | Coming soon|
| ExcelApi1.5  | Version 1703 (Build 8067.2070) or later| 2.2 or later |15.36 or later| March 2017 | Coming soon|
| ExcelApi1.4 | Version 1701 (Build 7870.2024) or later| 2.2 or later |15.36 or later| January 2017 | Coming soon|
| ExcelApi1.3  | Version 1608 (Build 7369.2055) or later| 1.27 or later |  15.27 or later| September 2016 | Version 1608 (Build 7601.6800) or later|
| ExcelApi1.2  | Version 1601 (Build 6741.2088) or later | 1.21 or later | 15.22 or later| January 2016 ||
| ExcelApi1.1  | Version 1509 (Build 4266.1001) or later | 1.19 or later | 15.20 or later| January 2016 ||

> **\*Note**: The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1 requirement set.

To find out more about versions, build numbers, and Office Online Server, see:

- [Version and build numbers of update channel releases for Office 365 clients](https://technet.microsoft.com/en-us/library/mt592918.aspx)
- [What version of Office am I using?](https://support.office.com/en-us/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19?ui=en-US&rs=en-US&ad=US&fromAR=1)
- [Where you can find the version and build number for an Office 365 client application](https://technet.microsoft.com/en-us/library/mt592918.aspx#Anchor_1)
- [Office Online Server overview](https://technet.microsoft.com/en-us/library/jj219437(v=office.16).aspx)

## Runtime requirement support check

During the runtime, add-ins can check if a particular host supports an API requirement set by doing the following-check: 

```js
if (Office.context.requirements.isSetSupported('ExcelApi', 1.3) === true) {
  /// perform actions
}
else {
  /// provide alternate flow/logic
}
```

## Manifest based requirement support check

Use the Requirements element in the add-in manifest to specify critical requirement sets or API members that your add-in must use. If the Office host or platform doesn't support the requirement sets or API members specified in the Requirements element, the add-in won't run in that host or platform, and won't display in My Add-ins. Instead, we recommend that you make your add-in available on all platforms of an Office host, such as Excel for Windows, Excel Online, and Excel for iPad. To make your add-in available on all Office hosts and platforms, use runtime checks instead of the Requirements element.

The following code example shows an add-in that loads in all Office host applications that support ExcelApi requirement set, version 1.3.

```xml
<Requirements>
   <Sets DefaultMinVersion="1.3">
      <Set Name="ExcelApi" MinVersion="1.3"/>
   </Sets>
</Requirements>
```

## What's new in Excel JavaScript 


|Object| What is new| Description|Feedback|
|:----|:----|:----|:----|
|[application](../excel/application.md)|_Property_ > calculationEngineVersion|Returns a number about the version of Excel Calculation Engine that the workbook was last fully recalculated by. Read-only.|beta|
|[application](../excel/application.md)|_Relationship_ > calculationState|Returns a CalculationState that indicates the calculation state of the application. Read-only.|beta|
|[application](../excel/application.md)|_Relationship_ > iterativeCalculation|Returns the Iterative Calculation settings. Read-only.|beta|
|[application](../excel/application.md)|_Method_ > [suspendScreenUpdatingUntilNextSync()]((../excel/application.md#suspendscreenupdatinguntilnextsync)|Suspends sceen updating until the next "context.sync()" is called.|beta|
|[autoFilter](../excel/autofilter.md)|_Property_ > enabled|Indicates if the AutoFilter is enabled or not. Read-Only. Read-only.|beta|
|[autoFilter](../excel/autofilter.md)|_Property_ > isDataFiltered|Indicates if the AutoFilter has filter criteria. Read-Only. Read-only.|beta|
|[autoFilter](../excel/autofilter.md)|_Relationship_ > criteria|Array that holds all filter criterias in an autofiltered range. Read-Only. Read-only.|beta|
|[autoFilter](../excel/autofilter.md)|_Method_ > [apply(range: Range or string, columnIndex: number, criteria: FilterCriteria)]((../excel/autofilter.md#applyrange-range-or-string-columnindex-number-criteria-filtercriteria)|Applies AutoFilter on a range and filters the column if column index and filter criteria are specified.|beta|
|[autoFilter](../excel/autofilter.md)|_Method_ > [clearCriteria()]((../excel/autofilter.md#clearcriteria)|Clears the criteria if AutoFilter has filters|beta|
|[autoFilter](../excel/autofilter.md)|_Method_ > [getRange()]((../excel/autofilter.md#getrange)|Returns the Range object that represents the range to which the AutoFilter applies.|beta|
|[autoFilter](../excel/autofilter.md)|_Method_ > [getRangeOrNullObject()]((../excel/autofilter.md#getrangeornullobject)|If there is Range object associated with the AutoFilter, this method returns it.|beta|
|[autoFilter](../excel/autofilter.md)|_Method_ > [reapply()]((../excel/autofilter.md#reapply)|Applies the specified Autofilter object currently on the range.|beta|
|[autoFilter](../excel/autofilter.md)|_Method_ > [remove()]((../excel/autofilter.md#remove)|Removes the AutoFilter for the range.|beta|
|[cellBorder](../excel/cellborder.md)|_Property_ > color|Specifies whether the match is case sensitive. Default is false (insensitive).|beta|
|[cellBorder](../excel/cellborder.md)|_Property_ > style|Specifies whether the match is case sensitive. Default is false (insensitive). Possible values are: None, Continuous, Dash, DashDot, DashDotDot, Dot, Double, SlantDashDot.|beta|
|[cellBorder](../excel/cellborder.md)|_Property_ > tintAndShade|Specifies whether the match is case sensitive. Default is false (insensitive). Possible values are: None, Continuous, Dash, DashDot, DashDotDot, Dot, Double, SlantDashDot.|beta|
|[cellBorder](../excel/cellborder.md)|_Property_ > weight|Specifies whether the match is case sensitive. Default is false (insensitive). Possible values are: None, Continuous, Dash, DashDot, DashDotDot, Dot, Double, SlantDashDot.|beta|
|[cellBorderCollection](../excel/cellbordercollection.md)|_Property_ > items|A collection of cellBorder objects. Read-only.|beta|
|[cellBorderCollection](../excel/cellbordercollection.md)|_Relationship_ > bottom|Specifies whether the match is case sensitive. Default is false (insensitive). Possible values are: None, Single, Double, SingleAccountant, DoubleAccountant.|beta|
|[cellBorderCollection](../excel/cellbordercollection.md)|_Relationship_ > diagonalDown|Specifies whether the match is case sensitive. Default is false (insensitive). Possible values are: None, Single, Double, SingleAccountant, DoubleAccountant.|beta|
|[cellBorderCollection](../excel/cellbordercollection.md)|_Relationship_ > diagonalUp|Specifies whether the match is case sensitive. Default is false (insensitive). Possible values are: None, Single, Double, SingleAccountant, DoubleAccountant.|beta|
|[cellBorderCollection](../excel/cellbordercollection.md)|_Relationship_ > horizontal|Specifies whether the match is case sensitive. Default is false (insensitive). Possible values are: None, Single, Double, SingleAccountant, DoubleAccountant.|beta|
|[cellBorderCollection](../excel/cellbordercollection.md)|_Relationship_ > left|Specifies whether the match is case sensitive. Default is false (insensitive). Possible values are: None, Single, Double, SingleAccountant, DoubleAccountant.|beta|
|[cellBorderCollection](../excel/cellbordercollection.md)|_Relationship_ > right|Specifies whether the match is case sensitive. Default is false (insensitive). Possible values are: None, Single, Double, SingleAccountant, DoubleAccountant.|beta|
|[cellBorderCollection](../excel/cellbordercollection.md)|_Relationship_ > top|Specifies whether the match is case sensitive. Default is false (insensitive). Possible values are: None, Single, Double, SingleAccountant, DoubleAccountant.|beta|
|[cellBorderCollection](../excel/cellbordercollection.md)|_Relationship_ > vertical|Specifies whether the match is case sensitive. Default is false (insensitive). Possible values are: None, Single, Double, SingleAccountant, DoubleAccountant.|beta|
|[cellPropertiesBorderLoadOptions](../excel/cellpropertiesborderloadoptions.md)|_Property_ > color|Specifies whether the match is case sensitive. Default is false (insensitive).|beta|
|[cellPropertiesBorderLoadOptions](../excel/cellpropertiesborderloadoptions.md)|_Property_ > style|Specifies whether the match is case sensitive. Default is false (insensitive).|beta|
|[cellPropertiesBorderLoadOptions](../excel/cellpropertiesborderloadoptions.md)|_Property_ > tintAndShade|Specifies whether the match is case sensitive. Default is false (insensitive).|beta|
|[cellPropertiesBorderLoadOptions](../excel/cellpropertiesborderloadoptions.md)|_Property_ > weight|Specifies whether the match is case sensitive. Default is false (insensitive).|beta|
|[cellPropertiesFill](../excel/cellpropertiesfill.md)|_Property_ > color|Specifies whether the match is case sensitive. Default is false (insensitive).|beta|
|[cellPropertiesFill](../excel/cellpropertiesfill.md)|_Property_ > patternColor|Specifies whether the match is case sensitive. Default is false (insensitive).|beta|
|[cellPropertiesFill](../excel/cellpropertiesfill.md)|_Property_ > patternTintAndShade|Specifies whether the match is case sensitive. Default is false (insensitive).|beta|
|[cellPropertiesFill](../excel/cellpropertiesfill.md)|_Property_ > tintAndShade|Specifies whether the match is case sensitive. Default is false (insensitive).|beta|
|[cellPropertiesFill](../excel/cellpropertiesfill.md)|_Relationship_ > pattern|Specifies whether the match is case sensitive. Default is false (insensitive).|beta|
|[cellPropertiesFillLoadOptions](../excel/cellpropertiesfillloadoptions.md)|_Property_ > color|Specifies whether the match is case sensitive. Default is false (insensitive).|beta|
|[cellPropertiesFillLoadOptions](../excel/cellpropertiesfillloadoptions.md)|_Property_ > pattern|Specifies whether the match is case sensitive. Default is false (insensitive).|beta|
|[cellPropertiesFillLoadOptions](../excel/cellpropertiesfillloadoptions.md)|_Property_ > patternColor|Specifies whether the match is case sensitive. Default is false (insensitive).|beta|
|[cellPropertiesFillLoadOptions](../excel/cellpropertiesfillloadoptions.md)|_Property_ > patternTintAndShade|Specifies whether the match is case sensitive. Default is false (insensitive).|beta|
|[cellPropertiesFillLoadOptions](../excel/cellpropertiesfillloadoptions.md)|_Property_ > tintAndShade|Specifies whether the match is case sensitive. Default is false (insensitive).|beta|
|[cellPropertiesFont](../excel/cellpropertiesfont.md)|_Property_ > bold|Specifies whether the match is case sensitive. Default is false (insensitive).|beta|
|[cellPropertiesFont](../excel/cellpropertiesfont.md)|_Property_ > color|Specifies whether the match is case sensitive. Default is false (insensitive).|beta|
|[cellPropertiesFont](../excel/cellpropertiesfont.md)|_Property_ > italic|Specifies whether the match is case sensitive. Default is false (insensitive).|beta|
|[cellPropertiesFont](../excel/cellpropertiesfont.md)|_Property_ > name|Specifies whether the match is case sensitive. Default is false (insensitive).|beta|
|[cellPropertiesFont](../excel/cellpropertiesfont.md)|_Property_ > size|Specifies whether the match is case sensitive. Default is false (insensitive).|beta|
|[cellPropertiesFont](../excel/cellpropertiesfont.md)|_Property_ > strikethrough|Specifies whether the match is case sensitive. Default is false (insensitive).|beta|
|[cellPropertiesFont](../excel/cellpropertiesfont.md)|_Property_ > subscript|Specifies whether the match is case sensitive. Default is false (insensitive).|beta|
|[cellPropertiesFont](../excel/cellpropertiesfont.md)|_Property_ > superscript|Specifies whether the match is case sensitive. Default is false (insensitive).|beta|
|[cellPropertiesFont](../excel/cellpropertiesfont.md)|_Property_ > tintAndShade|Specifies whether the match is case sensitive. Default is false (insensitive). Possible values are: None, Single, Double, SingleAccountant, DoubleAccountant.|beta|
|[cellPropertiesFont](../excel/cellpropertiesfont.md)|_Property_ > underline|Specifies whether the match is case sensitive. Default is false (insensitive). Possible values are: None, Single, Double, SingleAccountant, DoubleAccountant.|beta|
|[cellPropertiesFontLoadOptions](../excel/cellpropertiesfontloadoptions.md)|_Property_ > bold|Specifies whether the match is case sensitive. Default is false (insensitive).|beta|
|[cellPropertiesFontLoadOptions](../excel/cellpropertiesfontloadoptions.md)|_Property_ > color|Specifies whether the match is case sensitive. Default is false (insensitive).|beta|
|[cellPropertiesFontLoadOptions](../excel/cellpropertiesfontloadoptions.md)|_Property_ > italic|Specifies whether the match is case sensitive. Default is false (insensitive).|beta|
|[cellPropertiesFontLoadOptions](../excel/cellpropertiesfontloadoptions.md)|_Property_ > name|Specifies whether the match is case sensitive. Default is false (insensitive).|beta|
|[cellPropertiesFontLoadOptions](../excel/cellpropertiesfontloadoptions.md)|_Property_ > size|Specifies whether the match is case sensitive. Default is false (insensitive).|beta|
|[cellPropertiesFontLoadOptions](../excel/cellpropertiesfontloadoptions.md)|_Property_ > strikethrough|Specifies whether the match is case sensitive. Default is false (insensitive).|beta|
|[cellPropertiesFontLoadOptions](../excel/cellpropertiesfontloadoptions.md)|_Property_ > subscript|Specifies whether the match is case sensitive. Default is false (insensitive).|beta|
|[cellPropertiesFontLoadOptions](../excel/cellpropertiesfontloadoptions.md)|_Property_ > superscript|Specifies whether the match is case sensitive. Default is false (insensitive).|beta|
|[cellPropertiesFontLoadOptions](../excel/cellpropertiesfontloadoptions.md)|_Property_ > tintAndShade|Specifies whether the match is case sensitive. Default is false (insensitive).|beta|
|[cellPropertiesFontLoadOptions](../excel/cellpropertiesfontloadoptions.md)|_Property_ > underline|Specifies whether the match is case sensitive. Default is false (insensitive).|beta|
|[cellPropertiesProtection](../excel/cellpropertiesprotection.md)|_Property_ > formulaHidden|Specifies whether the match is case sensitive. Default is false (insensitive).|beta|
|[cellPropertiesProtection](../excel/cellpropertiesprotection.md)|_Property_ > locked|Specifies whether the match is case sensitive. Default is false (insensitive).|beta|
|[chart](../excel/chart.md)|_Relationship_ > pivotOptions|Encapsulates the options for the pivot chart. Read-only.|beta|
|[chart](../excel/chart.md)|_Method_ > [activate()]((../excel/chart.md#activate)|Activate the chart in the Excel UI.|beta|
|[chartAreaFormat](../excel/chartareaformat.md)|_Property_ > roundedCorners|TrueΓö¼├íif the chart area of the chart has rounded corners. ReadWrite.|beta|
|[chartAreaFormat](../excel/chartareaformat.md)|_Relationship_ > colorScheme|Returns or sets anΓö¼├íintegerΓö¼├íthat represents the color scheme for the chart. ReadWrite.|beta|
|[chartAxis](../excel/chartaxis.md)|_Property_ > linkNumberFormat|Represents whether the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells).|beta|
|[chartBinOptions](../excel/chartbinoptions.md)|_Property_ > allowOverflow|Returns or sets if bin overflow enabled in a histogram chart or pareto chart. ReadWrite.|beta|
|[chartBinOptions](../excel/chartbinoptions.md)|_Property_ > allowUnderflow|Returns or sets if bin underflow enabled in a histogram chart or pareto chart. ReadWrite.|beta|
|[chartBinOptions](../excel/chartbinoptions.md)|_Property_ > count|Returns or sets count of bin of a histogram chart or pareto chart. ReadWrite.|beta|
|[chartBinOptions](../excel/chartbinoptions.md)|_Property_ > overflowValue|Returns or sets bin overflow value of a histogram chart or pareto chart. ReadWrite.|beta|
|[chartBinOptions](../excel/chartbinoptions.md)|_Property_ > underflowValue|Returns or sets bin underflow value of a histogram chart or pareto chart. ReadWrite.|beta|
|[chartBinOptions](../excel/chartbinoptions.md)|_Property_ > width|Returns or sets bin width value of a histogram chart or pareto chart. ReadWrite.|beta|
|[chartBinOptions](../excel/chartbinoptions.md)|_Relationship_ > type|Returns or sets bin type of a histogram chart or pareto chart. ReadWrite.|beta|
|[chartBoxwhiskerOptions](../excel/chartboxwhiskeroptions.md)|_Property_ > showInnerPoints|Returns or sets if inner points showed in a Box &amp; whisker chart. ReadWrite.|beta|
|[chartBoxwhiskerOptions](../excel/chartboxwhiskeroptions.md)|_Property_ > showMeanLine|Returns or sets if mean line showed in a Box &amp; whisker chart. ReadWrite.|beta|
|[chartBoxwhiskerOptions](../excel/chartboxwhiskeroptions.md)|_Property_ > showMeanMarker|Returns or sets if mean marker showed in a Box &amp; whisker chart. ReadWrite.|beta|
|[chartBoxwhiskerOptions](../excel/chartboxwhiskeroptions.md)|_Property_ > showOutlierPoints|Returns or sets if outlier points showed in a Box &amp; whisker chart. ReadWrite.|beta|
|[chartBoxwhiskerOptions](../excel/chartboxwhiskeroptions.md)|_Relationship_ > quartileCalculation|Returns or sets quartile calculation type of a Box &amp; whisker chart. ReadWrite.|beta|
|[chartDataLabel](../excel/chartdatalabel.md)|_Property_ > linkNumberFormat|Boolean value representing if the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells).|beta|
|[chartDataLabels](../excel/chartdatalabels.md)|_Property_ > linkNumberFormat|Represents whether the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells).|beta|
|[chartErrorBars](../excel/charterrorbars.md)|_Property_ > endStyleCap|Represents whether have the end style cap for the error bars.|beta|
|[chartErrorBars](../excel/charterrorbars.md)|_Property_ > visible|Represents whether shown error bars.|beta|
|[chartErrorBars](../excel/charterrorbars.md)|_Relationship_ > format|Represents the formatting of chart ErrorBars. Read-only.|beta|
|[chartErrorBars](../excel/charterrorbars.md)|_Relationship_ > include|Represents which error-bar parts to include.|beta|
|[chartErrorBars](../excel/charterrorbars.md)|_Relationship_ > type|Represents the range marked by error bars.|beta|
|[chartErrorBarsFormat](../excel/charterrorbarsformat.md)|_Relationship_ > line|Represents chart line formatting. Read-only.|beta|
|[chartMapOptions](../excel/chartmapoptions.md)|_Relationship_ > labelStrategy|Returns or sets series map labels strategy of a region map chart. ReadWrite.|beta|
|[chartMapOptions](../excel/chartmapoptions.md)|_Relationship_ > level|Returns or sets series map area of a region map chart. ReadWrite.|beta|
|[chartMapOptions](../excel/chartmapoptions.md)|_Relationship_ > projectionType|Returns or sets series projection type of a region map chart. ReadWrite.|beta|
|[chartPivotOptions](../excel/chartpivotoptions.md)|_Property_ > showAxisFieldButtons|Represents whether to display axis field buttons on a PivotChart.|beta|
|[chartPivotOptions](../excel/chartpivotoptions.md)|_Property_ > showLegendFieldButtons|Represents whether to display legend field buttons on a PivotChart.|beta|
|[chartPivotOptions](../excel/chartpivotoptions.md)|_Property_ > showReportFilterFieldButtons|Represents whether to display report filter field buttons on a PivotChart.|beta|
|[chartPivotOptions](../excel/chartpivotoptions.md)|_Property_ > showValueFieldButtons|Represents whether to display show value field buttons on a PivotChart.|beta|
|[chartSeries](../excel/chartseries.md)|_Property_ > bubbleScale|Returns or sets the scale factor for bubbles in the specified chart group. Can be an integer value from 0 (zero) to 300, corresponding to a percentage of the default size. Applies only to bubble charts. ReadWrite.|beta|
|[chartSeries](../excel/chartseries.md)|_Property_ > gradientMaximumColor|Returns or sets the Color for maximum value of a region map chart series. ReadWrite.|beta|
|[chartSeries](../excel/chartseries.md)|_Property_ > gradientMaximumValue|Returns or sets the maximum value of a region map chart series. ReadWrite.|beta|
|[chartSeries](../excel/chartseries.md)|_Property_ > gradientMidpointColor|Returns or sets the Color for midpoint value of a region map chart series. ReadWrite.|beta|
|[chartSeries](../excel/chartseries.md)|_Property_ > gradientMidpointValue|Returns or sets the midpoint value of a region map chart series. ReadWrite.|beta|
|[chartSeries](../excel/chartseries.md)|_Property_ > gradientMinimumColor|Returns or sets the Color for minimum value of a region map chart series. ReadWrite.|beta|
|[chartSeries](../excel/chartseries.md)|_Property_ > gradientMinimumValue|Returns or sets the minimum value of a region map chart series. ReadWrite.|beta|
|[chartSeries](../excel/chartseries.md)|_Property_ > invertColor|Returns or sets the fill color for negative data points in a series. ReadWrite.|beta|
|[chartSeries](../excel/chartseries.md)|_Property_ > showConnectorLines|Returns or sets if connector lines show in a waterfall chart. ReadWrite.|beta|
|[chartSeries](../excel/chartseries.md)|_Property_ > showLeaderLines|True if Microsoft Excel show leaderlines for each datalabel in series. ReadWrite.|beta|
|[chartSeries](../excel/chartseries.md)|_Property_ > splitValue|Returns or sets the threshold value separating the two sections of either a pie of pie chart or a bar of pie chart. ReadWrite.|beta|
|[chartSeries](../excel/chartseries.md)|_Relationship_ > binOptions|Encapsulates the bin options only for histogram chart and pareto chart. Read-only.|beta|
|[chartSeries](../excel/chartseries.md)|_Relationship_ > boxwhiskerOptions|Encapsulates the options for the Box &amp; Whisker chart. Read-only.|beta|
|[chartSeries](../excel/chartseries.md)|_Relationship_ > gradientMaximumType|Returns or sets the type for maximum value of a region map chart series. ReadWrite.|beta|
|[chartSeries](../excel/chartseries.md)|_Relationship_ > gradientMidpointType|Returns or sets the type for midpoint value of a region map chart series. ReadWrite.|beta|
|[chartSeries](../excel/chartseries.md)|_Relationship_ > gradientMinimumType|Returns or sets the type for minimum value of a region map chart series. ReadWrite.|beta|
|[chartSeries](../excel/chartseries.md)|_Relationship_ > gradientStyle|Returns or sets series gradient style of a region map chart. ReadWrite.|beta|
|[chartSeries](../excel/chartseries.md)|_Relationship_ > mapOptions|Encapsulates the options for the Map chart. Read-only.|beta|
|[chartSeries](../excel/chartseries.md)|_Relationship_ > parentLabelStrategy|Returns or sets series parent label strategy area of a treemap chart. ReadWrite.|beta|
|[chartSeries](../excel/chartseries.md)|_Relationship_ > xErrorBars|Represents the error bar object for a chart series. Read-only.|beta|
|[chartSeries](../excel/chartseries.md)|_Relationship_ > yErrorBars|Represents the error bar object for a chart series. Read-only.|beta|
|[chartTrendlineLabel](../excel/charttrendlinelabel.md)|_Property_ > linkNumberFormat|Boolean value representing if the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells).|beta|
|[closeWorkbookPostProcessAction](../excel/closeworkbookpostprocessaction.md)|_Relationship_ > actionType|Transmits additional data to client side, e.g., worksheetId for TableSelectionChangedEvent. Possible values are: UpdateFilteredFunctionList, UpdateAllFunctionList, None.|beta|
|[comment](../excel/comment.md)|_Property_ > content|GetSet the content.|beta|
|[comment](../excel/comment.md)|_Property_ > id|Represents the comment identifier. Read-only.|beta|
|[comment](../excel/comment.md)|_Property_ > isParent|Represents whether it is a comment thread or reply. Always return true here. Read-only.|beta|
|[comment](../excel/comment.md)|_Relationship_ > replies|Represents a collection of reply objects associated with the comment. Read-only.|beta|
|[comment](../excel/comment.md)|_Method_ > [delete()]((../excel/comment.md#delete)|Deletes the comment thread.|beta|
|[commentCollection](../excel/commentcollection.md)|_Property_ > items|A collection of comment objects. Read-only.|beta|
|[commentCollection](../excel/commentcollection.md)|_Method_ > [add(content: string, cellAddress: object, contentType: ContentType)]((../excel/commentcollection.md#addcontent-string-celladdress-object-contenttype-contenttype)|Creates a new comment(comment thread) based on the cell location and content. Invalid argument will be thrown if the location is larger than one cell.|beta|
|[commentCollection](../excel/commentcollection.md)|_Method_ > [getCount()]((../excel/commentcollection.md#getcount)|Gets the number of comments in the collection.|beta|
|[commentCollection](../excel/commentcollection.md)|_Method_ > [getItem(commentId: string)]((../excel/commentcollection.md#getitemcommentid-string)|Returns a comment identified by its ID. Read-only.|beta|
|[commentCollection](../excel/commentcollection.md)|_Method_ > [getItemAt(index: number)]((../excel/commentcollection.md#getitematindex-number)|Gets a comment based on its position in the collection.|beta|
|[commentCollection](../excel/commentcollection.md)|_Method_ > [getItemByCell(cellAddress: object)]((../excel/commentcollection.md#getitembycellcelladdress-object)|Gets a comment on the specific cell in the collection.|beta|
|[commentCollection](../excel/commentcollection.md)|_Method_ > [getItemByReplyId(replyId: string)]((../excel/commentcollection.md#getitembyreplyidreplyid-string)|Gets a comment related to its reply ID in the collection.|beta|
|[commentReply](../excel/commentreply.md)|_Property_ > content|GetSet the content.|beta|
|[commentReply](../excel/commentreply.md)|_Property_ > id|Represents the comment reply identifier. Read-only.|beta|
|[commentReply](../excel/commentreply.md)|_Property_ > isParent|Represents whether it is a comment thread or reply. Always return false here. Read-only.|beta|
|[commentReply](../excel/commentreply.md)|_Method_ > [delete()]((../excel/commentreply.md#delete)|Deletes the comment reply.|beta|
|[commentReply](../excel/commentreply.md)|_Method_ > [getParentComment()]((../excel/commentreply.md#getparentcomment)|Get its parent comment of this reply.|beta|
|[commentReplyCollection](../excel/commentreplycollection.md)|_Property_ > items|A collection of commentReply objects. Read-only.|beta|
|[commentReplyCollection](../excel/commentreplycollection.md)|_Method_ > [add(content: string, contentType: ContentType)]((../excel/commentreplycollection.md#addcontent-string-contenttype-contenttype)|Creates a comment reply for comment.|beta|
|[commentReplyCollection](../excel/commentreplycollection.md)|_Method_ > [getCount()]((../excel/commentreplycollection.md#getcount)|Gets the number of comment replies in the collection.|beta|
|[commentReplyCollection](../excel/commentreplycollection.md)|_Method_ > [getItem(commentReplyId: string)]((../excel/commentreplycollection.md#getitemcommentreplyid-string)|Returns a comment reply identified by its ID. Read-only.|beta|
|[commentReplyCollection](../excel/commentreplycollection.md)|_Method_ > [getItemAt(index: number)]((../excel/commentreplycollection.md#getitematindex-number)|Gets a comment reply based on its position in the collection.|beta|
|[conditionalFormat](../excel/conditionalformat.md)|_Method_ > [getRanges()]((../excel/conditionalformat.md#getranges)|Returns the RangeAreas, comprising one or more rectangular ranges, the conditonal format is applied to. Read-only.|beta|
|[dataValidation](../excel/datavalidation.md)|_Method_ > [getInvalidCells()]((../excel/datavalidation.md#getinvalidcells)|Returns a RangeAreas, comprising one or more rectangular ranges, with invalid cell values. If all cell values are valid, this function will throw an ItemNotFound error.|beta|
|[dataValidation](../excel/datavalidation.md)|_Method_ > [getInvalidCellsOrNullObject()]((../excel/datavalidation.md#getinvalidcellsornullobject)|Returns a RangeAreas, comprising one or more rectangular ranges, with invalid cell values. If all cell values are valid, this function will return null.|beta|
|[filterCriteria](../excel/filtercriteria.md)|_Property_ > subField|The property used by the filter to do rich filter on richvalues.|beta|
|[geometricShape](../excel/geometricshape.md)|_Property_ > id|Represents the shape identifier. Read-only.|beta|
|[geometricShape](../excel/geometricshape.md)|_Relationship_ > shape|Returns the shape object for the geometric shape. Read-only.|beta|
|[groupShapeCollection](../excel/groupshapecollection.md)|_Property_ > items|A collection of groupShape objects. Read-only.|beta|
|[groupShapeCollection](../excel/groupshapecollection.md)|_Method_ > [getCount()]((../excel/groupshapecollection.md#getcount)|Returns the number of shapes in the group shape. Read-only.|beta|
|[groupShapeCollection](../excel/groupshapecollection.md)|_Method_ > [getItem(name: string)]((../excel/groupshapecollection.md#getitemname-string)|Gets a shape using its name.|beta|
|[groupShapeCollection](../excel/groupshapecollection.md)|_Method_ > [getItem(shapeId: string)]((../excel/groupshapecollection.md#getitemshapeid-string)|Returns a shape identified by the shape id. Read-only.|beta|
|[groupShapeCollection](../excel/groupshapecollection.md)|_Method_ > [getItemAt(index: number)]((../excel/groupshapecollection.md#getitematindex-number)|Gets a shape based on its position in the collection.|beta|
|[headerFooter](../excel/headerfooter.md)|_Property_ > centerFooter|Gets or sets the center footer of the worksheet.|beta|
|[headerFooter](../excel/headerfooter.md)|_Property_ > centerHeader|Gets or sets the center header of the worksheet.|beta|
|[headerFooter](../excel/headerfooter.md)|_Property_ > leftFooter|Gets or sets the left footer of the worksheet.|beta|
|[headerFooter](../excel/headerfooter.md)|_Property_ > leftHeader|Gets or sets the left header of the worksheet.|beta|
|[headerFooter](../excel/headerfooter.md)|_Property_ > rightFooter|Gets or sets the right footer of the worksheet.|beta|
|[headerFooter](../excel/headerfooter.md)|_Property_ > rightHeader|Gets or sets the right header of the worksheet.|beta|
|[headerFooterGroup](../excel/headerfootergroup.md)|_Property_ > useSheetMargins|Gets or sets a flag indicating if headersfooters are aligned with the page margins set in the page layout options for the worksheet.|beta|
|[headerFooterGroup](../excel/headerfootergroup.md)|_Property_ > useSheetScale|Gets or sets a flag indicating if headersfooters should be scaled by the page percentage scale set in the page layout options for the worksheet.|beta|
|[headerFooterGroup](../excel/headerfootergroup.md)|_Relationship_ > defaultForAllPages|The general headerfooter, used for all pages unless evenodd or first page is specified. Read-only.|beta|
|[headerFooterGroup](../excel/headerfootergroup.md)|_Relationship_ > evenPages|The headerfooter to use for even pages, odd headerfooter needs to be specified for odd pages. Read-only.|beta|
|[headerFooterGroup](../excel/headerfootergroup.md)|_Relationship_ > firstPage|The first page headerfooter, for all other pages general or evenodd is used. Read-only.|beta|
|[headerFooterGroup](../excel/headerfootergroup.md)|_Relationship_ > oddPages|The headerfooter to use for odd pages, even headerfooter needs to be specified for even pages. Read-only.|beta|
|[headerFooterGroup](../excel/headerfootergroup.md)|_Relationship_ > state|Gets or sets the state of which headersfooters are set.|beta|
|[image](../excel/image.md)|_Property_ > id|Represents the shape identifier for the image object. Read-only.|beta|
|[image](../excel/image.md)|_Relationship_ > format|Returns the format for the image. Read-only.|beta|
|[image](../excel/image.md)|_Relationship_ > shape|Returns the shape object for the image. Read-only.|beta|
|[iterativeCalculation](../excel/iterativecalculation.md)|_Property_ > enabled|True if Excel will use iteration to resolve circular references.|beta|
|[iterativeCalculation](../excel/iterativecalculation.md)|_Property_ > maxChange|Returns or sets the maximum amount of change between each iteration as Excel resolves circular references.|beta|
|[iterativeCalculation](../excel/iterativecalculation.md)|_Property_ > maxIteration|Returns or sets the maximum number of iterations that Excel can use to resolve a circular reference.|beta|
|[line](../excel/line.md)|_Property_ > id|Represents the shape identifier. Read-only.|beta|
|[line](../excel/line.md)|_Relationship_ > connectorType|Represents the connector type for the line.|beta|
|[line](../excel/line.md)|_Relationship_ > shape|Returns the shape object for the line. Read-only.|beta|
|[pageBreak](../excel/pagebreak.md)|_Property_ > columnIndex|Represents the column index for the page break Read-only.|beta|
|[pageBreak](../excel/pagebreak.md)|_Property_ > rowIndex|Represents the row index for the page break Read-only.|beta|
|[pageBreak](../excel/pagebreak.md)|_Method_ > [delete()]((../excel/pagebreak.md#delete)|Deletes a page break object.|beta|
|[pageBreak](../excel/pagebreak.md)|_Method_ > [getStartCell()]((../excel/pagebreak.md#getstartcell)|Gets the first cell after the page break.|beta|
|[pageBreakCollection](../excel/pagebreakcollection.md)|_Property_ > items|A collection of pageBreak objects. Read-only.|beta|
|[pageBreakCollection](../excel/pagebreakcollection.md)|_Method_ > [add(pageBreakRange: Range or string)]((../excel/pagebreakcollection.md#addpagebreakrange-range-or-string)|Adds a page break before the top-left cell of the range specified.|beta|
|[pageBreakCollection](../excel/pagebreakcollection.md)|_Method_ > [getCount()]((../excel/pagebreakcollection.md#getcount)|Gets the number of page breaks in the collection.|beta|
|[pageBreakCollection](../excel/pagebreakcollection.md)|_Method_ > [getItem(index: number)]((../excel/pagebreakcollection.md#getitemindex-number)|Gets a page break object via the index.|beta|
|[pageBreakCollection](../excel/pagebreakcollection.md)|_Method_ > [removePageBreaks()]((../excel/pagebreakcollection.md#removepagebreaks)|Resets all manual page breaks in the collection.|beta|
|[pageLayout](../excel/pagelayout.md)|_Property_ > blackAndWhite|Gets or sets the worksheet's black and white print option.|beta|
|[pageLayout](../excel/pagelayout.md)|_Property_ > bottomMargin|Gets or sets the worksheet's bottom page margin to use for printing in points.|beta|
|[pageLayout](../excel/pagelayout.md)|_Property_ > centerHorizontally|Gets or sets the worksheet's center horizontally flag. This flag determines whether the worksheet will be centered horizontally when it's printed.|beta|
|[pageLayout](../excel/pagelayout.md)|_Property_ > centerVertically|Gets or sets the worksheet's center vertically flag. This flag determines whether the worksheet will be centered vertically when it's printed.|beta|
|[pageLayout](../excel/pagelayout.md)|_Property_ > draftMode|Gets or sets the worksheet's draft mode option. If true the sheet will be printed without graphics.|beta|
|[pageLayout](../excel/pagelayout.md)|_Property_ > firstPageNumber|Gets or sets the worksheet's first page number to print. Null value represents "auto" page numbering.|beta|
|[pageLayout](../excel/pagelayout.md)|_Property_ > footerMargin|Gets or sets the worksheet's footer margin, in points, for use when printing.|beta|
|[pageLayout](../excel/pagelayout.md)|_Property_ > headerMargin|Gets or sets the worksheet's header margin, in points, for use when printing.|beta|
|[pageLayout](../excel/pagelayout.md)|_Property_ > leftMargin|Gets or sets the worksheet's left margin, in points, for use when printing.|beta|
|[pageLayout](../excel/pagelayout.md)|_Property_ > orientation|Gets or sets the worksheet's orientation of the page. Possible values are: Portrait, Landscape.|beta|
|[pageLayout](../excel/pagelayout.md)|_Property_ > paperSize|Gets or sets the worksheet's paper size of the page. Possible values are: Letter, LetterSmall, Tabloid, Ledger, Legal, Statement, Executive, A3, A4, A4Small, A5, B4, B5, Folio, Quatro, Paper10x14, Paper11x17, Note, Envelope9, Envelope10, Envelope11, Envelope12, Envelope14, Csheet, Dsheet, Esheet, EnvelopeDL, EnvelopeC5, EnvelopeC3, EnvelopeC4, EnvelopeC6, EnvelopeC65, EnvelopeB4, EnvelopeB5, EnvelopeB6, EnvelopeItaly, EnvelopeMonarch, EnvelopePersonal, FanfoldUS, FanfoldStdGerman, FanfoldLegalGerman.|beta|
|[pageLayout](../excel/pagelayout.md)|_Property_ > printGridlines|Gets or sets the worksheet's print gridlines flag. This flag determines whether gridlines will be printed or not.|beta|
|[pageLayout](../excel/pagelayout.md)|_Property_ > printHeadings|Gets or sets the worksheet's print headings flag. This flag determines whether headings will be printed or not.|beta|
|[pageLayout](../excel/pagelayout.md)|_Property_ > rightMargin|Gets or sets the worksheet's right margin, in points, for use when printing.|beta|
|[pageLayout](../excel/pagelayout.md)|_Property_ > topMargin|Gets or sets the worksheet's top margin, in points, for use when printing.|beta|
|[pageLayout](../excel/pagelayout.md)|_Relationship_ > headersFooters|Header and footer configuration for the worksheet. Read-only.|beta|
|[pageLayout](../excel/pagelayout.md)|_Relationship_ > printComments|Gets or sets whether the worksheet's comments should be displayed when printing.|beta|
|[pageLayout](../excel/pagelayout.md)|_Relationship_ > printErrors|Gets or sets the worksheet's print errors option.|beta|
|[pageLayout](../excel/pagelayout.md)|_Relationship_ > printOrder|Gets or sets the worksheet's page print order option. This specifies the order to use for processing the page number printed.|beta|
|[pageLayout](../excel/pagelayout.md)|_Relationship_ > zoom|Gets or sets the worksheet's print zoom options.|beta|
|[pageLayout](../excel/pagelayout.md)|_Method_ > [getPrintArea()]((../excel/pagelayout.md#getprintarea)|Gets the RangeAreas object, comprising one or more rectangular ranges, that represents the print area for the worksheet. If there is no print area, an ItemNotFound error will be thrown.|beta|
|[pageLayout](../excel/pagelayout.md)|_Method_ > [getPrintAreaOrNullObject()]((../excel/pagelayout.md#getprintareaornullobject)|Gets the RangeAreas object, comprising one or more rectangular ranges, that represents the print area for the worksheet. If there is no print area, a null object will be returned.|beta|
|[pageLayout](../excel/pagelayout.md)|_Method_ > [getPrintTitleColumns()]((../excel/pagelayout.md#getprinttitlecolumns)|Gets the range object representing the title columns.|beta|
|[pageLayout](../excel/pagelayout.md)|_Method_ > [getPrintTitleColumnsOrNullObject()]((../excel/pagelayout.md#getprinttitlecolumnsornullobject)|Gets the range object representing the title columns. If not set, this will return a null object.|beta|
|[pageLayout](../excel/pagelayout.md)|_Method_ > [getPrintTitleRows()]((../excel/pagelayout.md#getprinttitlerows)|Gets the range object representing the title rows.|beta|
|[pageLayout](../excel/pagelayout.md)|_Method_ > [getPrintTitleRowsOrNullObject()]((../excel/pagelayout.md#getprinttitlerowsornullobject)|Gets the range object representing the title rows. If not set, this will return a null object.|beta|
|[pageLayout](../excel/pagelayout.md)|_Method_ > [setPrintArea(printArea: Range or RangeAreas or string)]((../excel/pagelayout.md#setprintareaprintarea-range-or-rangeareas-or-string)|Sets the worksheet's print area.|beta|
|[pageLayout](../excel/pagelayout.md)|_Method_ > [setPrintMargins(unit: PrintMarginUnit, marginOptions: PageLayoutMarginOptions)]((../excel/pagelayout.md#setprintmarginsunit-printmarginunit-marginoptions-pagelayoutmarginoptions)|Sets the worksheet's page margins with units.|beta|
|[pageLayout](../excel/pagelayout.md)|_Method_ > [setPrintTitleColumns(printTitleColumns: Range or string)]((../excel/pagelayout.md#setprinttitlecolumnsprinttitlecolumns-range-or-string)|Sets the columns that contain the cells to be repeated at the left of each page of the worksheet for printing.|beta|
|[pageLayout](../excel/pagelayout.md)|_Method_ > [setPrintTitleRows(printTitleRows: Range or string)]((../excel/pagelayout.md#setprinttitlerowsprinttitlerows-range-or-string)|Sets the rows that contain the cells to be repeated at the top of each page of the worksheet for printing.|beta|
|[pageLayoutMarginOptions](../excel/pagelayoutmarginoptions.md)|_Property_ > bottom|Represents the page layout bottom margin in the unit specified to use for printing.|beta|
|[pageLayoutMarginOptions](../excel/pagelayoutmarginoptions.md)|_Property_ > footer|Represents the page layout footer margin in the unit specified to use for printing.|beta|
|[pageLayoutMarginOptions](../excel/pagelayoutmarginoptions.md)|_Property_ > header|Represents the page layout header margin in the unit specified to use for printing.|beta|
|[pageLayoutMarginOptions](../excel/pagelayoutmarginoptions.md)|_Property_ > left|Represents the page layout left margin in the unit specified to use for printing.|beta|
|[pageLayoutMarginOptions](../excel/pagelayoutmarginoptions.md)|_Property_ > right|Represents the page layout right margin in the unit specified to use for printing.|beta|
|[pageLayoutMarginOptions](../excel/pagelayoutmarginoptions.md)|_Property_ > top|Represents the page layout top margin in the unit specified to use for printing.|beta|
|[pageLayoutZoomOptions](../excel/pagelayoutzoomoptions.md)|_Property_ > horizontalFitToPages|Number of pages to fit horizontally. This value can be null if percentage scale is used.|beta|
|[pageLayoutZoomOptions](../excel/pagelayoutzoomoptions.md)|_Property_ > scale|Print page scale value can be between 10 and 400. This value can be null if fit to page tall or wide is specified.|beta|
|[pageLayoutZoomOptions](../excel/pagelayoutzoomoptions.md)|_Property_ > verticalFitToPages|Number of pages to fit vertically. This value can be null if percentage scale is used.|beta|
|[pivotField](../excel/pivotfield.md)|_Method_ > [sortByValues(sortby: SortBy, valuesHierarchy: DataPivotHierarchy, pivotItemScope: ()[])]((../excel/pivotfield.md#sortbyvaluessortby-sortby-valueshierarchy-datapivothierarchy-pivotitemscope-)|Sorts the PivotField by specified values in a given scope. The scope defines which specific values will be used to sort when|beta|
|[pivotLayout](../excel/pivotlayout.md)|_Property_ > autoFormat|True if formatting will be automatically formatted when it╬ô├ç├ûs refreshed or when fields are moved|beta|
|[pivotLayout](../excel/pivotlayout.md)|_Property_ > enableFieldList|True if the field list should be shown or hidden from the UI.|beta|
|[pivotLayout](../excel/pivotlayout.md)|_Property_ > preserveFormatting|True if formatting is preserved when the report is refreshed or recalculated by operations such as pivoting, sorting, or changing page field items.|beta|
|[pivotLayout](../excel/pivotlayout.md)|_Method_ > [getCell(dataHierarchy: DataPivotHierarchy or string, rowItems: ()[], columnItems: ()[])]((../excel/pivotlayout.md#getcelldatahierarchy-datapivothierarchy-or-string-rowitems--columnitems-)|Gets the cell in the PivotTable's data body that contains the value for the intersection of the specified dataHierarchy, rowItems, and columnItems.|beta|
|[pivotLayout](../excel/pivotlayout.md)|_Method_ > [getDataHierarchy(cell: Range or string)]((../excel/pivotlayout.md#getdatahierarchycell-range-or-string)|Gets the DataHierarchy that is used to calculate the value in a specified range within the PivotTable.|beta|
|[pivotLayout](../excel/pivotlayout.md)|_Method_ > [getPivotItems(axis: PivotAxis, cell: Range or string)]((../excel/pivotlayout.md#getpivotitemsaxis-pivotaxis-cell-range-or-string)|Gets the PivotItems from an axis that make up the value in a specified range within the PivotTable.|beta|
|[pivotLayout](../excel/pivotlayout.md)|_Method_ > [setAutosortOnCell(cell: Range or string, sortby: SortBy)]((../excel/pivotlayout.md#setautosortoncellcell-range-or-string-sortby-sortby)|Sets an autosort using the specified cell to automatically select all criteria and context for the sort.|beta|
|[pivotTable](../excel/pivottable.md)|_Property_ > enableDataValueEditing|True if the PivotTable should use custom lists when sorting.|beta|
|[pivotTable](../excel/pivottable.md)|_Property_ > useCustomSortLists|True if the PivotTable should use custom lists when sorting.|beta|
|[range](../excel/range.md)|_Property_ > hasSpill|Represents if all cells have a spill border. Read-only.|beta|
|[range](../excel/range.md)|_Relationship_ > linkedDataTypeState|Represents the data type state of each cell. Read-only.|beta|
|[range](../excel/range.md)|_Method_ > [autoFill(destinationRange: Range or string, autoFillType: AutoFillType)]((../excel/range.md#autofilldestinationrange-range-or-string-autofilltype-autofilltype)|Fills range from the current range to the destination range.|beta|
|[range](../excel/range.md)|_Method_ > [convertDataTypeToText()]((../excel/range.md#convertdatatypetotext)|Converts the range cells with datatypes into text.|beta|
|[range](../excel/range.md)|_Method_ > [convertToLinkedDataType(serviceID: number, languageCulture: string)]((../excel/range.md#converttolinkeddatatypeserviceid-number-languageculture-string)|Converts the range cells into linked datatype in the worksheet.|beta|
|[range](../excel/range.md)|_Method_ > [copyFrom(sourceRange: Range or RangeAreas or string, copyType: RangeCopyType, skipBlanks: bool, transpose: bool)]((../excel/range.md#copyfromsourcerange-range-or-rangeareas-or-string-copytype-rangecopytype-skipblanks-bool-transpose-bool)|Copies cell data or formatting from the source range or RangeAreas to the current range.|beta|
|[range](../excel/range.md)|_Method_ > [find(text: string, criteria: SearchCriteria)]((../excel/range.md#findtext-string-criteria-searchcriteria)|Finds the given string based on the criteria specified.|beta|
|[range](../excel/range.md)|_Method_ > [findOrNullObject(text: string, criteria: SearchCriteria)]((../excel/range.md#findornullobjecttext-string-criteria-searchcriteria)|Finds the given string based on the criteria specified.|beta|
|[range](../excel/range.md)|_Method_ > [getSpecialCells(cellType: SpecialCellType, cellValueType: SpecialCellValueType)]((../excel/range.md#getspecialcellscelltype-specialcelltype-cellvaluetype-specialcellvaluetype)|Gets the RangeAreas object, comprising one or more rectangular ranges, that represents all the cells that match the specified type and value.|beta|
|[range](../excel/range.md)|_Method_ > [getSpecialCellsOrNullObject(cellType: SpecialCellType, cellValueType: SpecialCellValueType)]((../excel/range.md#getspecialcellsornullobjectcelltype-specialcelltype-cellvaluetype-specialcellvaluetype)|Gets the RangeAreas object, comprising one or more ranges, that represents all the cells that match the specified type and value.|beta|
|[range](../excel/range.md)|_Method_ > [getSpillParent()]((../excel/range.md#getspillparent)|Gets the range object containing the anchor cell for a cell getting spilled into. Fails if applied to a range with more than one cell. Read only.|beta|
|[range](../excel/range.md)|_Method_ > [getSpillingToRange()]((../excel/range.md#getspillingtorange)|Gets the range object containing the spill range when called on an anchor cell. Fails if applied to a range with more than one cell. Read only.|beta|
|[range](../excel/range.md)|_Method_ > [getTables(fullyContained: bool)]((../excel/range.md#gettablesfullycontained-bool)|Gets a scoped collection of tables that overlap with the range.|beta|
|[range](../excel/range.md)|_Method_ > [removeDuplicates(columns: int[], includesHeader: bool)]((../excel/range.md#removeduplicatescolumns-int-includesheader-bool)|Removes duplicate values from the range specified by the columns.|beta|
|[range](../excel/range.md)|_Method_ > [replaceAll(text: string, replacement: string, criteria: ReplaceCriteria)]((../excel/range.md#replacealltext-string-replacement-string-criteria-replacecriteria)|Finds and replaces the given string based on the criteria specified within the current range.|beta|
|[range](../excel/range.md)|_Method_ > [setCellProperties(cellPropertiesData: CellPropertiesInternal[][])]((../excel/range.md#setcellpropertiescellpropertiesdata-cellpropertiesinternal)|Updates the range based on a 2D array of cell properties , encapsulating things like font, fill, borders, alignment, and so forth.|beta|
|[range](../excel/range.md)|_Method_ > [setColumnProperties(columnPropertiesData: CellPropertiesInternal[])]((../excel/range.md#setcolumnpropertiescolumnpropertiesdata-cellpropertiesinternal)|Updates the range based on a single-dimensional array of column properties, encapsulating things like font, fill, borders, alignment, and so forth.|beta|
|[range](../excel/range.md)|_Method_ > [setDirty()]((../excel/range.md#setdirty)|Set a range to be recalculated when the next recalculation occurs.|beta|
|[range](../excel/range.md)|_Method_ > [setRowProperties(rowPropertiesData: CellPropertiesInternal[])]((../excel/range.md#setrowpropertiesrowpropertiesdata-cellpropertiesinternal)|Updates the range based on a single-dimensional array of row properties, encapsulating things like font, fill, borders, alignment, and so forth.|beta|
|[rangeAreas](../excel/rangeareas.md)|_Property_ > address|Returns the RageAreas reference in A1-style. Address value will contain the worksheet name for each rectangular block of cells (e.g. "Sheet1!A1:B4, Sheet1!D1:D4"). Read-only.|beta|
|[rangeAreas](../excel/rangeareas.md)|_Property_ > addressLocal|Returns the RageAreas reference in the user locale. Read-only.|beta|
|[rangeAreas](../excel/rangeareas.md)|_Property_ > areaCount|Returns the number of rectangular ranges that comprise this RangeAreas object. Read-only.|beta|
|[rangeAreas](../excel/rangeareas.md)|_Property_ > cellCount|Returns the number of cells in the RangeAreas object, summing up the cell counts of all of the individual rectangular ranges. Returns -1 if the cell count exceeds 2^31-1 (2,147,483,647). Read-only.|beta|
|[rangeAreas](../excel/rangeareas.md)|_Property_ > isEntireColumn|Indicates whether all the ranges on this RangeAreas object represent entire columns (e.g., "A:C, Q:Z"). Read-only.|beta|
|[rangeAreas](../excel/rangeareas.md)|_Property_ > isEntireRow|Indicates whether all the ranges on this RangeAreas object represent entire rows (e.g., "1:3, 5:7"). Read-only.|beta|
|[rangeAreas](../excel/rangeareas.md)|_Property_ > style|Represents the style for all ranges in this RangeAreas object.|beta|
|[rangeAreas](../excel/rangeareas.md)|_Relationship_ > areas|Returns a collection of rectangular ranges that comprise this RangeAreas object. Read-only.|beta|
|[rangeAreas](../excel/rangeareas.md)|_Relationship_ > conditionalFormats|Returns a collection of ConditionalFormats that intersect with any cells in this RangeAreas object. Read-only.|beta|
|[rangeAreas](../excel/rangeareas.md)|_Relationship_ > dataValidation|Returns a dataValidation object for all ranges in the RangeAreas. Read-only.|beta|
|[rangeAreas](../excel/rangeareas.md)|_Relationship_ > format|Returns a rangeFormat object, encapsulating the the font, fill, borders, alignment, and other properties for all ranges in the RangeAreas object. Read-only.|beta|
|[rangeAreas](../excel/rangeareas.md)|_Relationship_ > worksheet|Returns the worksheet for the current RangeAreas. Read-only.|beta|
|[rangeAreas](../excel/rangeareas.md)|_Method_ > [calculate()]((../excel/rangeareas.md#calculate)|Calculates all cells in the RangeAreas.|beta|
|[rangeAreas](../excel/rangeareas.md)|_Method_ > [clear(applyTo: string)]((../excel/rangeareas.md#clearapplyto-string)|Clears values, format, fill, border, etc on each of the areas that comprise this RangeAreas object.|beta|
|[rangeAreas](../excel/rangeareas.md)|_Method_ > [convertDataTypeToText()]((../excel/rangeareas.md#convertdatatypetotext)|Converts all cells in the RangeAreas with datatypes into text.|beta|
|[rangeAreas](../excel/rangeareas.md)|_Method_ > [convertToLinkedDataType(serviceID: number, languageCulture: string)]((../excel/rangeareas.md#converttolinkeddatatypeserviceid-number-languageculture-string)|Converts all cells in the RangeAreas into linked datatype.|beta|
|[rangeAreas](../excel/rangeareas.md)|_Method_ > [copyFrom(sourceRange: Range or RangeAreas or string, copyType: RangeCopyType, skipBlanks: bool, transpose: bool)]((../excel/rangeareas.md#copyfromsourcerange-range-or-rangeareas-or-string-copytype-rangecopytype-skipblanks-bool-transpose-bool)|Copies cell data or formatting from the source range or RangeAreas to the current RangeAreas.|beta|
|[rangeAreas](../excel/rangeareas.md)|_Method_ > [getEntireColumn()]((../excel/rangeareas.md#getentirecolumn)|Returns a RangeAreas object that represents the entire columns of the RangeAreas (for example, if the current RangeAreas represents cells "B4:E11, H2", it returns a RangeAreas that represents columns "B:E, H:H").|beta|
|[rangeAreas](../excel/rangeareas.md)|_Method_ > [getEntireRow()]((../excel/rangeareas.md#getentirerow)|Returns a RangeAreas object that represents the entire rows of the RangeAreas (for example, if the current RangeAreas represents cells "B4:E11", it returns a RangeAreas that represents rows "4:11").|beta|
|[rangeAreas](../excel/rangeareas.md)|_Method_ > [getIntersection(anotherRange: Range or RangeAreas or string)]((../excel/rangeareas.md#getintersectionanotherrange-range-or-rangeareas-or-string)|Returns the RangeAreas object that represents the intersection of the given ranges or RangeAreas. If no intersection is found, an ItemNotFound error will be thrown.|beta|
|[rangeAreas](../excel/rangeareas.md)|_Method_ > [getIntersectionOrNullObject(anotherRange: Range or RangeAreas or string)]((../excel/rangeareas.md#getintersectionornullobjectanotherrange-range-or-rangeareas-or-string)|Returns the RangeAreas object that represents the intersection of the given ranges or RangeAreas. If no intersection is found, a null object is returned.|beta|
|[rangeAreas](../excel/rangeareas.md)|_Method_ > [getOffsetRangeAreas(rowOffset: number, columnOffset: number)]((../excel/rangeareas.md#getoffsetrangeareasrowoffset-number-columnoffset-number)|Returns an RangeAreas object that is shifted by the specific row and column offset. The dimension of the returned RangeAreas will match the original object. If the resulting RangeAreas is forced outside the bounds of the worksheet grid, an error will be thrown.|beta|
|[rangeAreas](../excel/rangeareas.md)|_Method_ > [getSpecialCells(cellType: SpecialCellType, cellValueType: SpecialCellValueType)]((../excel/rangeareas.md#getspecialcellscelltype-specialcelltype-cellvaluetype-specialcellvaluetype)|Returns a RangeAreas object that represents all the cells that match the specified type and value. Throws an error if no special cells are found that match the criteria.|beta|
|[rangeAreas](../excel/rangeareas.md)|_Method_ > [getSpecialCellsOrNullObject(cellType: SpecialCellType, cellValueType: SpecialCellValueType)]((../excel/rangeareas.md#getspecialcellsornullobjectcelltype-specialcelltype-cellvaluetype-specialcellvaluetype)|Returns a RangeAreas object that represents all the cells that match the specified type and value. Returns a null object if no special cells are found that match the criteria.|beta|
|[rangeAreas](../excel/rangeareas.md)|_Method_ > [getTables(fullyContained: bool)]((../excel/rangeareas.md#gettablesfullycontained-bool)|Returns a scoped collection of tables that overlap with any range in this RangeAreas object.|beta|
|[rangeAreas](../excel/rangeareas.md)|_Method_ > [getUsedRangeAreas(valuesOnly: bool)]((../excel/rangeareas.md#getusedrangeareasvaluesonly-bool)|Returns the used RangeAreas that comprises all the used areas of individual rectangular ranges in the RangeAreas object.|beta|
|[rangeAreas](../excel/rangeareas.md)|_Method_ > [getUsedRangeAreasOrNullObject(valuesOnly: bool)]((../excel/rangeareas.md#getusedrangeareasornullobjectvaluesonly-bool)|Returns the used RangeAreas that comprises all the used areas of individual rectangular ranges in the RangeAreas object.|beta|
|[rangeAreas](../excel/rangeareas.md)|_Method_ > [setDirty()]((../excel/rangeareas.md#setdirty)|Sets the RangeAreas to be recalculated when the next recalculation occurs.|beta|
|[rangeBorder](../excel/rangeborder.md)|_Property_ > tintAndShade|Returns or sets a double that lightens or darkens a color for Range Border, the value is between -1 (darkest) and 1 (brightest), with 0 for the original color.|beta|
|[rangeBorderCollection](../excel/rangebordercollection.md)|_Property_ > tintAndShade|Returns or sets a double that lightens or darkens a color for Range Borders, the value is between -1 (darkest) and 1 (brightest), with 0 for the original color.|beta|
|[rangeCollection](../excel/rangecollection.md)|_Property_ > items|A collection of range objects. Read-only.|beta|
|[rangeCollection](../excel/rangecollection.md)|_Method_ > [getCount()]((../excel/rangecollection.md#getcount)|Returns the number of ranges in the RangeCollection.|beta|
|[rangeCollection](../excel/rangecollection.md)|_Method_ > [getItemAt(index: number)]((../excel/rangecollection.md#getitematindex-number)|Returns the range object based on its position in the RangeCollection.|beta|
|[rangeFill](../excel/rangefill.md)|_Property_ > patternColor|Sets HTML color code representing the color of the Range pattern, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|beta|
|[rangeFill](../excel/rangefill.md)|_Property_ > patternTintAndShade|Returns or sets a double that lightens or darkens a pattern color for Range Fill, the value is between -1 (darkest) and 1 (brightest), with 0 for the original color.|beta|
|[rangeFill](../excel/rangefill.md)|_Property_ > tintAndShade|Returns or sets a double that lightens or darkens a color for Range Fill, the value is between -1 (darkest) and 1 (brightest), with 0 for the original color.|beta|
|[rangeFill](../excel/rangefill.md)|_Relationship_ > pattern|Gets or sets the pattern of a Range.|beta|
|[rangeFont](../excel/rangefont.md)|_Property_ > strikethrough|Represents the strikethrough status of font. A null value indicates that the entire range doesn't have uniform Strikethrough setting.|beta|
|[rangeFont](../excel/rangefont.md)|_Property_ > subscript|Represents the Subscript status of font.|beta|
|[rangeFont](../excel/rangefont.md)|_Property_ > superscript|Represents the Superscript status of font.|beta|
|[rangeFont](../excel/rangefont.md)|_Property_ > tintAndShade|Returns or sets a double that lightens or darkens a color for Range Font, the value is between -1 (darkest) and 1 (brightest), with 0 for the original color.|beta|
|[rangeFormat](../excel/rangeformat.md)|_Property_ > autoIndent|Indicates if text is automatically indented when text alignment is set to equal distribution.|beta|
|[rangeFormat](../excel/rangeformat.md)|_Property_ > indentLevel|An integer from 0 to 250 that indicates the indent level.|beta|
|[rangeFormat](../excel/rangeformat.md)|_Property_ > readingOrder|The reading order for the range. Possible values are: Context, LeftToRight, RightToLeft.|beta|
|[rangeFormat](../excel/rangeformat.md)|_Property_ > shrinkToFit|Indicates if text automatically shrinks to fit in the available column width.|beta|
|[removeDuplicatesResult](../excel/removeduplicatesresult.md)|_Property_ > removed|Number of duplicated rows removed by the operation. Read-only.|beta|
|[removeDuplicatesResult](../excel/removeduplicatesresult.md)|_Property_ > uniqueRemaining|Number of remaining unique rows present in the resulting range. Read-only.|beta|
|[replaceCriteria](../excel/replacecriteria.md)|_Property_ > completeMatch|Specifies whether the match needs to be complete or partial. Default is false (partial).|beta|
|[replaceCriteria](../excel/replacecriteria.md)|_Property_ > matchCase|Specifies whether the match is case sensitive. Default is false (insensitive).|beta|
|[searchCriteria](../excel/searchcriteria.md)|_Property_ > completeMatch|Specifies whether the match needs to be complete or partial. Default is false (partial).|beta|
|[searchCriteria](../excel/searchcriteria.md)|_Property_ > matchCase|Specifies whether the match is case sensitive. Default is false (insensitive).|beta|
|[searchCriteria](../excel/searchcriteria.md)|_Relationship_ > searchDirection|Specifies the search direction. Default is forward.|beta|
|[shape](../excel/shape.md)|_Property_ > altTextDescription|Returns or sets the alternative descriptive text string for a Shape object when the object is saved to a Web page.|beta|
|[shape](../excel/shape.md)|_Property_ > altTextTitle|Returns or sets the alternative title text string for a Shape object when the object is saved to a Web page.|beta|
|[shape](../excel/shape.md)|_Property_ > height|Represents the height, in points, of the shape.|beta|
|[shape](../excel/shape.md)|_Property_ > id|Represents the shape identifier. Read-only.|beta|
|[shape](../excel/shape.md)|_Property_ > left|The distance, in points, from the left side of the shape to the left of the worksheet.|beta|
|[shape](../excel/shape.md)|_Property_ > level{|Represents the level of the specified shape. Level 0 means the shape is not part of any group, level 1 means the shape is part of a top-level group, etc. Read-only.|beta|
|[shape](../excel/shape.md)|_Property_ > lockAspectRatio|Represents if the aspect ratio locked, in boolean, of the shape.|beta|
|[shape](../excel/shape.md)|_Property_ > name|Represents the name of the shape.|beta|
|[shape](../excel/shape.md)|_Property_ > rotation|Represents the rotation, in degrees, of the shape.|beta|
|[shape](../excel/shape.md)|_Property_ > top|The distance, in points, from the top edge of the shape to the top of the worksheet.|beta|
|[shape](../excel/shape.md)|_Property_ > visible{|Represents the visibility, in boolean, of the specified shape.|beta|
|[shape](../excel/shape.md)|_Property_ > width|Represents the width, in points, of the shape.|beta|
|[shape](../excel/shape.md)|_Property_ > zOrderPosition|Returns the position of the specified shape in the z-order, the very bottom shape's z-order value is 0. Read-only.|beta|
|[shape](../excel/shape.md)|_Relationship_ > fill|Returns the fill formatting of the shape object. Read-only.|beta|
|[shape](../excel/shape.md)|_Relationship_ > geometricShape|Returns the geometric shape for the shape object. Error will be thrown, if the shape object is other shape type (Like, Image, SmartArt, etc.) rather than GeometricShape. Read-only.|beta|
|[shape](../excel/shape.md)|_Relationship_ > geometricShapeType|Represents the geometric shape type of the specified shape.|beta|
|[shape](../excel/shape.md)|_Relationship_ > group|Returns the shape group for the shape object. Error will be thrown, if the shape object is other shape type (Like, Image, SmartArt, etc.) rather than GroupShape. Read-only.|beta|
|[shape](../excel/shape.md)|_Relationship_ > image|Returns the image for the shape object. Error will be thrown, if the shape object is other shape type (Like, GeometricShape, SmartArt, etc.) rather than Image. Read-only.|beta|
|[shape](../excel/shape.md)|_Relationship_ > line|Returns the line object for the shape object. Error will be thrown, if the shape object is other shape type (Like, GeometricShape, SmartArt, etc.) rather than Image. Read-only.|beta|
|[shape](../excel/shape.md)|_Relationship_ > lineFormat|Returns the line formatting of the shape object. Read-only.|beta|
|[shape](../excel/shape.md)|_Relationship_ > parentGroup{|Represents the parent group of the specified shape. Read-only.|beta|
|[shape](../excel/shape.md)|_Relationship_ > placement|Represents the placment, value that represents the way the object is attached to the cells below it.|beta|
|[shape](../excel/shape.md)|_Relationship_ > textFrame|Returns the textFrame object of a shape. Read only. Read-only.|beta|
|[shape](../excel/shape.md)|_Relationship_ > type|Returns the type of the specified shape. Read-only.|beta|
|[shape](../excel/shape.md)|_Method_ > [delete()]((../excel/shape.md#delete)|Deletes the Shape|beta|
|[shape](../excel/shape.md)|_Method_ > [incrementLeft(increment: double)]((../excel/shape.md#incrementleftincrement-double)|Moves the shape horizontally by the specified number of points.|beta|
|[shape](../excel/shape.md)|_Method_ > [incrementRotation(increment: double)]((../excel/shape.md#incrementrotationincrement-double)|Changes the rotation of the shape around the z-axis by the specified number of degrees.|beta|
|[shape](../excel/shape.md)|_Method_ > [incrementTop(increment: double)]((../excel/shape.md#incrementtopincrement-double)|Moves the shape vertically by the specified number of points.|beta|
|[shape](../excel/shape.md)|_Method_ > [saveAsPicture(format: PictureFormat)]((../excel/shape.md#saveaspictureformat-pictureformat)|Saves the shape as a picture and returns the picture in the form of base64 encoded string, using the DPI sets to 96. Only support saves as to Excel.PictureFormat.BMP, Excel.PictureFormat.PNG, Excel.PictureFormat.JPEG and Excel.PictureFormat.GIF.|beta|
|[shape](../excel/shape.md)|_Method_ > [scaleHeight(scaleFactor: double, scaleType: ShapeScaleType, scaleFrom: ShapeScaleFrom)]((../excel/shape.md#scaleheightscalefactor-double-scaletype-shapescaletype-scalefrom-shapescalefrom)|Scales the height of the shape by a specified factor. For pictures, you can indicate whether you want to scale the shape relative to the original or the current size. Shapes other than pictures are always scaled relative to their current height.|beta|
|[shape](../excel/shape.md)|_Method_ > [scaleWidth(scaleFactor: double, scaleType: ShapeScaleType, scaleFrom: ShapeScaleFrom)]((../excel/shape.md#scalewidthscalefactor-double-scaletype-shapescaletype-scalefrom-shapescalefrom)|Scales the width of the shape by a specified factor. For pictures, you can indicate whether you want to scale the shape relative to the original or the current size. Shapes other than pictures are always scaled relative to their current width.|beta|
|[shape](../excel/shape.md)|_Method_ > [setZOrder(value: string)]((../excel/shape.md#setzordervalue-string)|Moves the specified shape in front of or behind other shapes in the collection (that is, changes the shape's position in the z-order).|beta|
|[shapeActivatedEventArgs](../excel/shapeactivatedeventargs.md)|_Property_ > shapeId|Gets the id of the shape that is activated.|beta|
|[shapeActivatedEventArgs](../excel/shapeactivatedeventargs.md)|_Property_ > type|Gets the type of the event.|beta|
|[shapeActivatedEventArgs](../excel/shapeactivatedeventargs.md)|_Property_ > worksheetId|Gets the id of the worksheet in which the shape is activated.|beta|
|[shapeCollection](../excel/shapecollection.md)|_Property_ > items|A collection of shape objects. Read-only.|beta|
|[shapeCollection](../excel/shapecollection.md)|_Method_ > [addGeometricShape(geometricShapeType: string, left: double, top: double, width: double, height: double)]((../excel/shapecollection.md#addgeometricshapegeometricshapetype-string-left-double-top-double-width-double-height-double)|Adds a geometric shape to worksheet. Returns a Shape object that represents the new shape.|beta|
|[shapeCollection](../excel/shapecollection.md)|_Method_ > [addGroup(values: ()[])]((../excel/shapecollection.md#addgroupvalues-)|Group a subset of shapes in a worksheet. Returns a Shape object that represents the new group of shapes.|beta|
|[shapeCollection](../excel/shapecollection.md)|_Method_ > [addImage(base64ImageString: string)]((../excel/shapecollection.md#addimagebase64imagestring-string)|Creates an image from a base64 string and adds it to worksheet. Returns the Shape object that represents the new Image.|beta|
|[shapeCollection](../excel/shapecollection.md)|_Method_ > [addLine(startLeft: double, startTop: double, endLeft: double, endTop: double, connectorType: string)]((../excel/shapecollection.md#addlinestartleft-double-starttop-double-endleft-double-endtop-double-connectortype-string)|Adds a line to worksheet. Returns a Shape object that represents the new line.|beta|
|[shapeCollection](../excel/shapecollection.md)|_Method_ > [addSVG(xmlImageString: string)]((../excel/shapecollection.md#addsvgxmlimagestring-string)|Creates an SVG from a XML string and adds it to worksheet. Returns a Shape object that represents the new Image.|beta|
|[shapeCollection](../excel/shapecollection.md)|_Method_ > [addTextBox(text: string)]((../excel/shapecollection.md#addtextboxtext-string)|Adds a textbox to worksheet by telling it's text content. Returns a Shape object that represents the new text box.|beta|
|[shapeCollection](../excel/shapecollection.md)|_Method_ > [getCount()]((../excel/shapecollection.md#getcount)|Returns the number of shapes in the worksheet. Read-only.|beta|
|[shapeCollection](../excel/shapecollection.md)|_Method_ > [getItem(shapeId: string)]((../excel/shapecollection.md#getitemshapeid-string)|Returns a shape identified by the shape id. Read-only.|beta|
|[shapeCollection](../excel/shapecollection.md)|_Method_ > [getItem(name: string)]((../excel/shapecollection.md#getitemname-string)|Gets a shape using its name.|beta|
|[shapeCollection](../excel/shapecollection.md)|_Method_ > [getItemAt(index: number)]((../excel/shapecollection.md#getitematindex-number)|Gets a shape based on its position in the collection.|beta|
|[shapeDeactivatedEventArgs](../excel/shapedeactivatedeventargs.md)|_Property_ > shapeId|Gets the id of the shape that is deactivated.|beta|
|[shapeDeactivatedEventArgs](../excel/shapedeactivatedeventargs.md)|_Property_ > type|Gets the type of the event.|beta|
|[shapeDeactivatedEventArgs](../excel/shapedeactivatedeventargs.md)|_Property_ > worksheetId|Gets the id of the worksheet in which the shape is deactivated.|beta|
|[shapeFill](../excel/shapefill.md)|_Property_ > foreColor|Represents the shape fill fore color in HTML color format, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange")|beta|
|[shapeFill](../excel/shapefill.md)|_Property_ > transparency|Returns or sets the degree of transparency of the specified fill as a value from 0.0 (opaque) through 1.0 (clear). For API not supported shape types  or special fill type with inconsistent transparencies, return null. For example, gradient fill type could have inconsistent transparencies.|beta|
|[shapeFill](../excel/shapefill.md)|_Relationship_ > type|Returns the fill type of the shape. Read-only.|beta|
|[shapeFill](../excel/shapefill.md)|_Method_ > [clear()]((../excel/shapefill.md#clear)|Clears the fill formatting of a shape object.|beta|
|[shapeFill](../excel/shapefill.md)|_Method_ > [setSolidColor(color: string)]((../excel/shapefill.md#setsolidcolorcolor-string)|Sets the fill formatting of a shape object to a uniform color, fill type changeing to Solid Fill.|beta|
|[shapeFont](../excel/shapefont.md)|_Property_ > bold|Represents the bold status of font. Returns null the TextRange includes both bold and non-bold text fragments.|beta|
|[shapeFont](../excel/shapefont.md)|_Property_ > color|HTML color code representation of the text color. E.g. #FF0000 represents Red. Returns null if the TextRange includes text fragments with different colors.|beta|
|[shapeFont](../excel/shapefont.md)|_Property_ > italic|Represents the italic status of font. Return null if the TextRange includes both italic and non-italic text fragments.|beta|
|[shapeFont](../excel/shapefont.md)|_Property_ > name|Represents font name (e.g. "Calibri"). If the text is Complex Script or East Asian language, represents corresponding font name; otherwise represents Latin font name.|beta|
|[shapeFont](../excel/shapefont.md)|_Property_ > size|Represents font size in points (e.g. 11). Return null if the TextRange includes text fragments with different font sizes.|beta|
|[shapeFont](../excel/shapefont.md)|_Relationship_ > underline|Type of underline applied to the font. Return null if the TextRange includes text fragments with different underline styles.|beta|
|[shapeGroup](../excel/shapegroup.md)|_Property_ > id|Represents the shape identifier. Read-only.|beta|
|[shapeGroup](../excel/shapegroup.md)|_Relationship_ > shape|Returns the shape object for the group. Read-only.|beta|
|[shapeGroup](../excel/shapegroup.md)|_Relationship_ > shapes|Returns the shape collection in the group. Read-only.|beta|
|[shapeGroup](../excel/shapegroup.md)|_Method_ > [ungroup()]((../excel/shapegroup.md#ungroup)|Ungroups any grouped shapes in the specified shape group.|beta|
|[shapeLineFormat](../excel/shapelineformat.md)|_Property_ > color|Represents the line color in HTML color format, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|beta|
|[shapeLineFormat](../excel/shapelineformat.md)|_Property_ > transparency|Represents the degree of transparency of the specified line as a value from 0.0 (opaque) through 1.0 (clear). Returns null when the shape has mixed line transparency property (e.g. group type of shape).|beta|
|[shapeLineFormat](../excel/shapelineformat.md)|_Property_ > visible|Represents whether the line formatting of a shape element is visible. Returns null when the shape has mixed line visible property (e.g. group type of shape).|beta|
|[shapeLineFormat](../excel/shapelineformat.md)|_Property_ > weight|Represents weight of the line, in points. Returns null when the line is not visible or has mixed line weight property (e.g. group type of shape).|beta|
|[shapeLineFormat](../excel/shapelineformat.md)|_Relationship_ > dashStyle|Represents the line style of the shape. Returns null when line is not visible or has mixed line dash style property (e.g. group type of shape).|beta|
|[shapeLineFormat](../excel/shapelineformat.md)|_Relationship_ > style|Represents the line style of the shape object. Returns null when line is not visible or has mixed line visible property (e.g. group type of shape).|beta|
|[slicer](../excel/slicer.md)|_Property_ > caption|Represents the caption of slicer.|beta|
|[slicer](../excel/slicer.md)|_Property_ > height|Represents the shape, in points, of the slicer.|beta|
|[slicer](../excel/slicer.md)|_Property_ > id|Represents the unique id of slicer. Read-only.|beta|
|[slicer](../excel/slicer.md)|_Property_ > isFilterCleared|True if all filters currently applied on the slicer is cleared. Read-only.|beta|
|[slicer](../excel/slicer.md)|_Property_ > left|Represents the distance, in points, from the left side of the slicer to the left of the worksheet.|beta|
|[slicer](../excel/slicer.md)|_Property_ > name|Represents the name of slicer.|beta|
|[slicer](../excel/slicer.md)|_Property_ > nameInFormula|Represents the name used in the formula.|beta|
|[slicer](../excel/slicer.md)|_Property_ > style|Constant value that represents the Slicer style. Possible values are: SlicerStyleLight1 thru SlicerStyleLight6, TableStyleOther1 thru TableStyleOther2, SlicerStyleDark1 thru SlicerStyleDark6. A custom user-defined style present in the workbook can also be specified.|beta|
|[slicer](../excel/slicer.md)|_Property_ > top|Represents the distance, in points, from the top edge of the slicer to the right of the worksheet.|beta|
|[slicer](../excel/slicer.md)|_Property_ > width|Represents the width, in points, of the slicer.|beta|
|[slicer](../excel/slicer.md)|_Relationship_ > slicerItems|Represents the collection of SlicerItems that are part of the slicer. Read-only.|beta|
|[slicer](../excel/slicer.md)|_Relationship_ > sortBy|Represents the sort order of the items in the slicer.|beta|
|[slicer](../excel/slicer.md)|_Relationship_ > worksheet|Represents the worksheet containing the slicer. Read-only.|beta|
|[slicer](../excel/slicer.md)|_Method_ > [clearFilters()]((../excel/slicer.md#clearfilters)|Clears all the filters currently applied on the slicer.|beta|
|[slicer](../excel/slicer.md)|_Method_ > [delete()]((../excel/slicer.md#delete)|Deletes the slicer.|beta|
|[slicer](../excel/slicer.md)|_Method_ > [getSelectedItems()]((../excel/slicer.md#getselecteditems)|Returns an array of selected items' names. Read-only.|beta|
|[slicer](../excel/slicer.md)|_Method_ > [selectItems(items: string[])]((../excel/slicer.md#selectitemsitems-string)|Select slicer items based on their names. Previous selection will be cleared.|beta|
|[slicerCollection](../excel/slicercollection.md)|_Property_ > items|A collection of slicer objects. Read-only.|beta|
|[slicerCollection](../excel/slicercollection.md)|_Method_ > [add(slicerSource: object, sourceField: object, slicerDestination: object)]((../excel/slicercollection.md#addslicersource-object-sourcefield-object-slicerdestination-object)|Adds a new slicer to the workbook.|beta|
|[slicerCollection](../excel/slicercollection.md)|_Method_ > [getCount()]((../excel/slicercollection.md#getcount)|Returns the number of slicers in the collection.|beta|
|[slicerCollection](../excel/slicercollection.md)|_Method_ > [getItem(key: string)]((../excel/slicercollection.md#getitemkey-string)|Gets a slicer object using its name or id.|beta|
|[slicerCollection](../excel/slicercollection.md)|_Method_ > [getItemAt(index: number)]((../excel/slicercollection.md#getitematindex-number)|Gets a slicer based on its position in the collection.|beta|
|[slicerCollection](../excel/slicercollection.md)|_Method_ > [getItemOrNullObject(key: string)]((../excel/slicercollection.md#getitemornullobjectkey-string)|Gets a slicer using its name or id. If the slicer does not exist, will return a null object.|beta|
|[slicerItem](../excel/sliceritem.md)|_Property_ > hasData|True if the slicer item has data. Read-only.|beta|
|[slicerItem](../excel/sliceritem.md)|_Property_ > isSelected|True if the slicer item is selected. Setting this value will not clear other SlicerItems' selected state.|beta|
|[slicerItem](../excel/sliceritem.md)|_Property_ > key|Represents the unique value representing the slicer item. Read-only.|beta|
|[slicerItem](../excel/sliceritem.md)|_Property_ > name|Represents the value displayed on UI. Read-only.|beta|
|[slicerItemCollection](../excel/sliceritemcollection.md)|_Property_ > items|A collection of slicerItem objects. Read-only.|beta|
|[slicerItemCollection](../excel/sliceritemcollection.md)|_Method_ > [getCount()]((../excel/sliceritemcollection.md#getcount)|Returns the number of slicer items in the slicer.|beta|
|[slicerItemCollection](../excel/sliceritemcollection.md)|_Method_ > [getItem(key: string)]((../excel/sliceritemcollection.md#getitemkey-string)|Gets a slicer item object using its key or name.|beta|
|[slicerItemCollection](../excel/sliceritemcollection.md)|_Method_ > [getItemAt(index: number)]((../excel/sliceritemcollection.md#getitematindex-number)|Gets a slicer item based on its position in the collection.|beta|
|[slicerItemCollection](../excel/sliceritemcollection.md)|_Method_ > [getItemOrNullObject(key: string)]((../excel/sliceritemcollection.md#getitemornullobjectkey-string)|Gets a slicer item using its key or name. If the slicer item does not exist, will return a null object.|beta|
|[sortField](../excel/sortfield.md)|_Property_ > subField|Represents the subfield that is the target property name of a rich value to sort on.|beta|
|[table](../excel/table.md)|_Relationship_ > autoFilter|Represents the AutoFilter object of the table. Read-Only. Read-only.|beta|
|[table](../excel/table.md)|_Method_ > [clearStyle()]((../excel/table.md#clearstyle)|Changes the table to use the default table style.|beta|
|[tableAddedEventArgs](../excel/tableaddedeventargs.md)|_Property_ > tableId|Gets the id of the table that is added.|beta|
|[tableAddedEventArgs](../excel/tableaddedeventargs.md)|_Property_ > type|Gets the type of the event.|beta|
|[tableAddedEventArgs](../excel/tableaddedeventargs.md)|_Property_ > worksheetId|Gets the id of the worksheet in which the table is added.|beta|
|[tableAddedEventArgs](../excel/tableaddedeventargs.md)|_Relationship_ > source|Gets the source of the event.|beta|
|[tableDeletedEventArgs](../excel/tabledeletedeventargs.md)|_Property_ > tableId|Specifies the id of the table that is deleted.|beta|
|[tableDeletedEventArgs](../excel/tabledeletedeventargs.md)|_Property_ > tableName|Specifies the name of the table that is deleted.|beta|
|[tableDeletedEventArgs](../excel/tabledeletedeventargs.md)|_Property_ > type|Specifies the type of the event.|beta|
|[tableDeletedEventArgs](../excel/tabledeletedeventargs.md)|_Property_ > worksheetId|Specifies the id of the worksheet in which the table is deleted.|beta|
|[tableDeletedEventArgs](../excel/tabledeletedeventargs.md)|_Relationship_ > source|Specifies the source of the event.|beta|
|[tableFilteredEventArgs](../excel/tablefilteredeventargs.md)|_Property_ > tableId|Represents the id of the table in which the filter is applied..|beta|
|[tableFilteredEventArgs](../excel/tablefilteredeventargs.md)|_Property_ > type|Represents the type of the event.|beta|
|[tableFilteredEventArgs](../excel/tablefilteredeventargs.md)|_Property_ > worksheetId|Represents the id of the worksheet which contains the table.|beta|
|[tableScopedCollection](../excel/tablescopedcollection.md)|_Property_ > items|A collection of tableScoped objects. Read-only.|beta|
|[tableScopedCollection](../excel/tablescopedcollection.md)|_Method_ > [getCount()]((../excel/tablescopedcollection.md#getcount)|Gets the number of tables in the collection.|beta|
|[tableScopedCollection](../excel/tablescopedcollection.md)|_Method_ > [getFirst()]((../excel/tablescopedcollection.md#getfirst)|Gets the first table in the collection. The tables in the collection are sorted top to bottom and left to right, such that top left table is the first table in the collection.|beta|
|[tableScopedCollection](../excel/tablescopedcollection.md)|_Method_ > [getItem(key: string)]((../excel/tablescopedcollection.md#getitemkey-string)|Gets a table by Name or ID.|beta|
|[textFrame](../excel/textframe.md)|_Property_ > bottomMargin|Represents the bottom margin, in points, of the text frame.|beta|
|[textFrame](../excel/textframe.md)|_Property_ > hasText|Specifies whether the TextFrame contains text. Read-only.|beta|
|[textFrame](../excel/textframe.md)|_Property_ > leftMargin|Represents the left margin, in points, of the text frame.|beta|
|[textFrame](../excel/textframe.md)|_Property_ > rightMargin|Represents the right margin, in points, of the text frame.|beta|
|[textFrame](../excel/textframe.md)|_Property_ > topMargin|Represents the top margin, in points, of the text frame.|beta|
|[textFrame](../excel/textframe.md)|_Relationship_ > autoSize|Gets or sets the auto sizing settings for the text frame. A text frame can be set to auto size the text to fit the text frame, or auto size the text frame to fit the text, or without auto sizing.|beta|
|[textFrame](../excel/textframe.md)|_Relationship_ > horizontalAlignment|Represents the horizontal alignment of the text frame.|beta|
|[textFrame](../excel/textframe.md)|_Relationship_ > horizontalOverflow|Represents the horizontal overflow type of the text frame.|beta|
|[textFrame](../excel/textframe.md)|_Relationship_ > orientation|Represents the text orientation of the text frame.|beta|
|[textFrame](../excel/textframe.md)|_Relationship_ > readingOrder|Represents the reading order of the text frame, RTL or LTR.|beta|
|[textFrame](../excel/textframe.md)|_Relationship_ > textRange|Represents the degree of transparency of the specified line as a value from 0.0 (opaque) through 1.0 (clear). Returns null when the shape has mixed line transparency property (e.g. group type of shape). Read-only.|beta|
|[textFrame](../excel/textframe.md)|_Relationship_ > verticalAlignment|Represents the vertical alignment of the text frame.|beta|
|[textFrame](../excel/textframe.md)|_Relationship_ > verticalOverflow|Represents the vertical overflow type of the text frame.|beta|
|[textFrame](../excel/textframe.md)|_Method_ > [deleteText()]((../excel/textframe.md#deletetext)|Deletes all the text in the textframe.|beta|
|[textRange](../excel/textrange.md)|_Property_ > text|Represents the plain text content of the text range.|beta|
|[textRange](../excel/textrange.md)|_Relationship_ > font|Returns a ShapeFont object that represents the font attributes for the text range. Read-only.|beta|
|[textRange](../excel/textrange.md)|_Method_ > [getCharacters(start: number, length: number)]((../excel/textrange.md#getcharactersstart-number-length-number)|Returns a TextRange object for characters in the given range.|beta|
|[workbook](../excel/workbook.md)|_Property_ > autoSave|True if the workbook is in auto save mode. Read-only.|beta|
|[workbook](../excel/workbook.md)|_Property_ > calculationEngineVersion|Returns a number about the version of Excel Calculation Engine. Read-Only. Read-only.|beta|
|[workbook](../excel/workbook.md)|_Property_ > chartDataPointTrack|True if all charts in the workbook are tracking the actual data points to which they are attached.|beta|
|[workbook](../excel/workbook.md)|_Property_ > isDirty|True if no changes have been made to the specified workbook since it was last saved.|beta|
|[workbook](../excel/workbook.md)|_Property_ > previouslySaved|True if the workbook has ever been saved locally or online. Read-only.|beta|
|[workbook](../excel/workbook.md)|_Property_ > use1904DateSystem|True if the workbook uses the 1904 date system.|beta|
|[workbook](../excel/workbook.md)|_Property_ > usePrecisionAsDisplayed|True if calculations in this workbook will be done using only the precision of the numbers as they're displayed.|beta|
|[workbook](../excel/workbook.md)|_Relationship_ > comments|Represents a collection of Comments associated with the workbook. Read-only.|beta|
|[workbook](../excel/workbook.md)|_Relationship_ > slicers|Represents a collection of Slicers associated with the workbook. Read-only.|beta|
|[workbook](../excel/workbook.md)|_Method_ > [close(closeBehavior: CloseBehavior)]((../excel/workbook.md#closeclosebehavior-closebehavior)|Close current workbook.|beta|
|[workbook](../excel/workbook.md)|_Method_ > [getActiveChart()]((../excel/workbook.md#getactivechart)|Gets the currently active chart in the workbook. If there is no active chart, will throw exception when invoke this statement|beta|
|[workbook](../excel/workbook.md)|_Method_ > [getActiveChartOrNullObject()]((../excel/workbook.md#getactivechartornullobject)|Gets the currently active chart in the workbook. If there is no active chart, will return null object|beta|
|[workbook](../excel/workbook.md)|_Method_ > [getActiveSlicer()]((../excel/workbook.md#getactiveslicer)|Gets the currently active slicer in the workbook. If there is no active slicer, will throw exception when invoke this statement.|beta|
|[workbook](../excel/workbook.md)|_Method_ > [getActiveSlicerOrNullObject()]((../excel/workbook.md#getactiveslicerornullobject)|Gets the currently active slicer in the workbook. If there is no active slicer, will return null object|beta|
|[workbook](../excel/workbook.md)|_Method_ > [getIsActiveCollabSession()]((../excel/workbook.md#getisactivecollabsession)|True if the workbook is being edited by multiple users (co-authoring).|beta|
|[workbook](../excel/workbook.md)|_Method_ > [getSelectedRanges()]((../excel/workbook.md#getselectedranges)|Gets the currently selected one or more ranges from the workbook. Unlike getSelectedRange(), this method returns a RangeAreas object that represents all the selected ranges.|beta|
|[workbook](../excel/workbook.md)|_Method_ > [save(saveBehavior: SaveBehavior)]((../excel/workbook.md#savesavebehavior-savebehavior)|Save current workbook.|beta|
|[workbookAutoSaveSettingChangedEventArgs](../excel/workbookautosavesettingchangedeventargs.md)|_Property_ > type|Represents the type of the event.|beta|
|[worksheet](../excel/worksheet.md)|_Property_ > enableCalculation|Gets or sets the enableCalculation property of the worksheet.|beta|
|[worksheet](../excel/worksheet.md)|_Relationship_ > autoFilter|Represents the AutoFilter object of the worksheet. Read-Only. Read-only.|beta|
|[worksheet](../excel/worksheet.md)|_Relationship_ > comments|Returns a collection of all the Comments objects on the worksheet. Read-only.|beta|
|[worksheet](../excel/worksheet.md)|_Relationship_ > horizontalPageBreaks|Gets the horizontal page break collection for the worksheet. This collection only contains manual page breaks. Read-only.|beta|
|[worksheet](../excel/worksheet.md)|_Relationship_ > pageLayout|Gets the PageLayout object of the worksheet. Read-only.|beta|
|[worksheet](../excel/worksheet.md)|_Relationship_ > shapes|Returns the collection of all the Shape objects on the worksheet. Read-only.|beta|
|[worksheet](../excel/worksheet.md)|_Relationship_ > slicers|Returns collection of slicers that are part of the worksheet. Read-only.|beta|
|[worksheet](../excel/worksheet.md)|_Relationship_ > verticalPageBreaks|Gets the vertical page break collection for the worksheet. This collection only contains manual page breaks. Read-only.|beta|
|[worksheet](../excel/worksheet.md)|_Method_ > [findAll(text: string, criteria: WorksheetSearchCriteria)]((../excel/worksheet.md#findalltext-string-criteria-worksheetsearchcriteria)|Finds all occurrences of the given string based on the criteria specified and returns them as a RangeAreas object, comprising one or more rectangular ranges.|beta|
|[worksheet](../excel/worksheet.md)|_Method_ > [findAllOrNullObject(text: string, criteria: WorksheetSearchCriteria)]((../excel/worksheet.md#findallornullobjecttext-string-criteria-worksheetsearchcriteria)|Finds all occurrences of the given string based on the criteria specified and returns them as a RangeAreas object, comprising one or more rectangular ranges.|beta|
|[worksheet](../excel/worksheet.md)|_Method_ > [getRanges(address: string)]((../excel/worksheet.md#getrangesaddress-string)|Gets the RangeAreas object, representing one or more blocks of rectangular ranges, specified by the address or name.|beta|
|[worksheet](../excel/worksheet.md)|_Method_ > [replaceAll(text: string, replacement: string, criteria: ReplaceCriteria)]((../excel/worksheet.md#replacealltext-string-replacement-string-criteria-replacecriteria)|Finds and replaces the given string based on the criteria specified within the current worksheet.|beta|
|[worksheetCollection](../excel/worksheetcollection.md)|_Method_ > [addFromBase64(base64File: string, sheetNamesToInsert: string[], positionType: WorksheetPositionType, relativeTo: Worksheet or string)]((../excel/worksheetcollection.md#addfrombase64base64file-string-sheetnamestoinsert-string-positiontype-worksheetpositiontype-relativeto-worksheet-or-string)|Inserts the specified worksheets of a workbook into the current workbook.|beta|
|[worksheetFilteredEventArgs](../excel/worksheetfilteredeventargs.md)|_Property_ > type|Represents the type of the event.|beta|
|[worksheetFilteredEventArgs](../excel/worksheetfilteredeventargs.md)|_Property_ > worksheetId|Represents the id of the worksheet in which the filter is applied.|beta|
|[worksheetFormatChangedEventArgs](../excel/worksheetformatchangedeventargs.md)|_Property_ > address|Gets the range address that represents the changed area of a specific worksheet.|beta|
|[worksheetFormatChangedEventArgs](../excel/worksheetformatchangedeventargs.md)|_Property_ > type|Gets the type of the event.|beta|
|[worksheetFormatChangedEventArgs](../excel/worksheetformatchangedeventargs.md)|_Property_ > worksheetId|Gets the id of the worksheet in which the data changed.|beta|
|[worksheetFormatChangedEventArgs](../excel/worksheetformatchangedeventargs.md)|_Relationship_ > source|Gets the source of the event.|beta|
|[worksheetSearchCriteria](../excel/worksheetsearchcriteria.md)|_Property_ > completeMatch|Specifies whether the match needs to be complete or partial. Default is false (partial).|beta|
|[worksheetSearchCriteria](../excel/worksheetsearchcriteria.md)|_Property_ > matchCase|Specifies whether the match is case sensitive. Default is false (insensitive).|beta|
## Additional resources

- [Specify Office hosts and API requirements](../../docs/overview/specify-office-hosts-and-api-requirements.md)
- [Office Add-ins XML manifest](../../docs/overview/add-in-manifests.md)
