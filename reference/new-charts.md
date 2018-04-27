|Object| What is new| Description|Req. Set|
|:----|:----|:----|:----|
|[chart](../markdown/chart.md)|_Property_ > categoryLabelLevel|Returns or sets a ChartCategoryLabelLevel enumeration constant referring to the level of where the category labels are being sourced from. ReadWrite.|beta|
|[chart](../markdown/chart.md)|_Property_ > colorScheme|Returns or sets an integer that represents the color scheme for the chart. ReadWrite.|beta|
|[chart](../markdown/chart.md)|_Property_ > displayBlanksAs|Returns or sets the way that blank cells are plotted on a chart. ReadWrite. Possible values are: NotPlotted, Zero, Interplotted.|beta|
|[chart](../markdown/chart.md)|_Property_ > plotBy|Returns or sets the way columns or rows are used as data series on the chart. ReadWrite. Possible values are: Rows, Columns.|beta|
|[chart](../markdown/chart.md)|_Property_ > plotVisibleOnly|True if only visible cells are plotted. False if both visible and hidden cells are plotted. ReadWrite.|beta|
|[chart](../markdown/chart.md)|_Property_ > roundedCorners|True if the chart area of the chart has rounded corners. ReadWrite.|beta|
|[chart](../markdown/chart.md)|_Property_ > seriesNameLevel|Returns or sets a ChartSeriesNameLevel enumeration constant referring to|beta|
|[chart](../markdown/chart.md)|_Property_ > showAxisFieldButtons|Represents whether to display axis field buttons on a PivotChart. The ShowAxisFieldButtons property corresponds to the Show Axis Field Buttons command on the Field Buttons drop-down list of the Analyze tab, which is available when a PivotChart is selected.|beta|
|[chart](../markdown/chart.md)|_Property_ > showDataLabelsOverMaximum|Represents whether to to show the data labels when the value is greater than the maximum value on the value axis. If value axis became smaller than the size of data points, you can use this property to set whether to show the data labels. This property applies to 2-D charts only.|beta|
|[chart](../markdown/chart.md)|_Property_ > showLegendFieldButtons|Represents whether to display legend field buttons on a PivotChart.|beta|
|[chart](../markdown/chart.md)|_Property_ > showReportFilterFieldButtons|Represents whether to display report filter field buttons on a PivotChart.|beta|
|[chart](../markdown/chart.md)|_Property_ > showValueFieldButtons|Represents whether to display show value field buttons on a PivotChart.|beta|
|[chart](../markdown/chart.md)|_Property_ > style|Returns or sets the chart style for the chart. ReadWrite.|beta|
|[chart](../markdown/chart.md)|_Relationship_ > plotArea|Represents the plotArea for the chart. Read-only.|beta|
|[chartAxis](../markdown/chartaxis.md)|_Property_ > isBetweenCategories|Represents whether value axis crosses the category axis between categories.|beta|
|[chartAxis](../markdown/chartaxis.md)|_Property_ > multiLevel|Represents whether an axis is multilevel or not.|beta|
|[chartAxis](../markdown/chartaxis.md)|_Property_ > numberFormat|Represents the format code for the axis tick label.|beta|
|[chartAxis](../markdown/chartaxis.md)|_Property_ > numberFormatLinked|Represents whether the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells).|beta|
|[chartAxis](../markdown/chartaxis.md)|_Property_ > offset|Represents the distance between the levels of labels, and the distance between the first level and the axis line. The value should be an integer from 0 to 1000.|beta|
|[chartAxis](../markdown/chartaxis.md)|_Property_ > positionAt|Represents the specified axis position where the other axis crosses at. Read Only. Set to this property should use SetPositionAt(double) method. Read-only.|beta|
|[chartAxis](../markdown/chartaxis.md)|_Property_ > textOrientation|Represents the text orientation of the axis tick label. The value should be an integer either from -90 to 90, or 180 for vertically-oriented text.|beta|
|[chartAxis](../markdown/chartaxis.md)|_Relationship_ > alignment|Represents the alignment for the specified axis tick label. Possible values are: Center, Left, Right, Justify, Distributed.|beta|
|[chartAxis](../markdown/chartaxis.md)|_Relationship_ > position|Represents the specified axis position where the other axis crosses.|beta|
|[chartAxis](../markdown/chartaxis.md)|_Method_ > [setPositionAt(value: double)](../markdown/chartaxis.md#setpositionatvalue-double)|Set the specified axis position where the other axis crosses at.|beta|
|[chartAxisFormat](../markdown/chartaxisformat.md)|_Relationship_ > fill|Represents chart fill formatting. Read-only.|beta|
|[chartAxisTitle](../markdown/chartaxistitle.md)|_Method_ > [setFormula(formula: string)](../markdown/chartaxistitle.md#setformulaformula-string)|A string value that represents the formula of chart axis title using A1-style notation.|beta|
|[chartAxisTitleFormat](../markdown/chartaxistitleformat.md)|_Relationship_ > border|Represents the border format, which includes color, linestyle and weight. Read-only.|beta|
|[chartAxisTitleFormat](../markdown/chartaxistitleformat.md)|_Relationship_ > fill|Represents chart fill formatting. Read-only.|beta|
|[chartBorder](../markdown/chartborder.md)|_Method_ > [clear()](../markdown/chartborder.md#clear)|Clear the border format of a chart element.|beta|
|[chartDataLabel](../markdown/chartdatalabel.md)|_Property_ > autoText|Boolean value representing if data label automatically generates appropriate text based on context.|beta|
|[chartDataLabel](../markdown/chartdatalabel.md)|_Property_ > formula|String value that represents the formula of chart data label using A1-style notation.|beta|
|[chartDataLabel](../markdown/chartdatalabel.md)|_Property_ > height|Returns the height, in points, of the chart data label. Read-only. Null if chart data label is not visible. Read-only.|beta|
|[chartDataLabel](../markdown/chartdatalabel.md)|_Property_ > horizontalAlignment|Represents the horizontal alignment for chart data label. Possible values are: Center, Left, Right, Justify, Distributed.|beta|
|[chartDataLabel](../markdown/chartdatalabel.md)|_Property_ > left|Represents the distance, in points, from the left edge of chart data label to the left edge of chart area. Null if chart data label is not visible.|beta|
|[chartDataLabel](../markdown/chartdatalabel.md)|_Property_ > numberFormat|String value that represents the format code for data label.|beta|
|[chartDataLabel](../markdown/chartdatalabel.md)|_Property_ > numberFormatLinked|Boolean value representing if the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells).|beta|
|[chartDataLabel](../markdown/chartdatalabel.md)|_Property_ > text|String representing the text of the data label on a chart.|beta|
|[chartDataLabel](../markdown/chartdatalabel.md)|_Property_ > textOrientation|Represents the text orientation of chart data label. The value should be an integer either from -90 to 90, or 180 for vertically-oriented text.|beta|
|[chartDataLabel](../markdown/chartdatalabel.md)|_Property_ > top|Represents the distance, in points, from the top edge of chart data label to the top of chart area. Null if chart data label is not visible.|beta|
|[chartDataLabel](../markdown/chartdatalabel.md)|_Property_ > verticalAlignment|Represents the vertical alignment of chart data label. Possible values are: Center, Bottom, Top, Justify, Distributed.|beta|
|[chartDataLabel](../markdown/chartdatalabel.md)|_Property_ > width|Returns the width, in points, of the chart data label. Read-only. Null if chart data label is not visible. Read-only.|beta|
|[chartDataLabel](../markdown/chartdatalabel.md)|_Relationship_ > format|Represents the format of chart data label. Read-only.|beta|
|[chartDataLabelFormat](../markdown/chartdatalabelformat.md)|_Relationship_ > border|Represents the border format, which includes color, linestyle and weight. Read-only.|beta|
|[chartDataLabels](../markdown/chartdatalabels.md)|_Property_ > autoText|Represents whether data labels automatically generates appropriate text based on context.|beta|
|[chartDataLabels](../markdown/chartdatalabels.md)|_Property_ > horizontalAlignment|Represents the horizontal alignment for chart data label. Possible values are: Center, Left, Right, Justify, Distributed.|beta|
|[chartDataLabels](../markdown/chartdatalabels.md)|_Property_ > numberFormat|Represents the format code for data labels.|beta|
|[chartDataLabels](../markdown/chartdatalabels.md)|_Property_ > numberFormatLinked|Represents whether the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells).|beta|
|[chartDataLabels](../markdown/chartdatalabels.md)|_Property_ > textOrientation|Represents the text orientation of data labels. The value should be an integer either from -90 to 90, or 180 for vertically-oriented text.|beta|
|[chartDataLabels](../markdown/chartdatalabels.md)|_Property_ > verticalAlignment|Represents the vertical alignment of chart data label. Possible values are: Center, Bottom, Top, Justify, Distributed.|beta|
|[chartErrorBars](../markdown/charterrorbars.md)|_Property_ > endStyleCap|Represents whether have the end style cap for the error bars.|beta|
|[chartErrorBars](../markdown/charterrorbars.md)|_Property_ > visible|Represents whether shown error bars.|beta|
|[chartErrorBars](../markdown/charterrorbars.md)|_Relationship_ > format|Represents the formatting of chart ErrorBars. Read-only.|beta|
|[chartErrorBars](../markdown/charterrorbars.md)|_Relationship_ > include|Represents which error-bar parts to include.|beta|
|[chartErrorBars](../markdown/charterrorbars.md)|_Relationship_ > type|Represents the range marked by error bars.|beta|
|[chartErrorBarsFormat](../markdown/charterrorbarsformat.md)|_Relationship_ > line|Represents chart line formatting. Read-only.|beta|
|[chartFormatString](../markdown/chartformatstring.md)|_Relationship_ > font|Represents the font attributes, such as font name, font size, color, etc. of chart characters object. Read-only.|beta|
|[chartLegendEntry](../markdown/chartlegendentry.md)|_Property_ > height|Represents the height of the legendEntry on the chart Legend. Read-only.|beta|
|[chartLegendEntry](../markdown/chartlegendentry.md)|_Property_ > index|Represents the index of the LegendEntry in the Chart Legend. Read-only.|beta|
|[chartLegendEntry](../markdown/chartlegendentry.md)|_Property_ > left|Represents the left of a chart legendEntry. Read-only.|beta|
|[chartLegendEntry](../markdown/chartlegendentry.md)|_Property_ > top|Represents the top of a chart legendEntry. Read-only.|beta|
|[chartLegendEntry](../markdown/chartlegendentry.md)|_Property_ > width|Represents the width of the legendEntry on the chart Legend. Read-only.|beta|
|[chartLegendFormat](../markdown/chartlegendformat.md)|_Relationship_ > border|Represents the border format, which includes color, linestyle and weight. Read-only.|beta|
|[chartPlotArea](../markdown/chartplotarea.md)|_Property_ > height|Represents the height value of plotArea.|beta|
|[chartPlotArea](../markdown/chartplotarea.md)|_Property_ > insideHeight|Represents the insideHeight value of plotArea.|beta|
|[chartPlotArea](../markdown/chartplotarea.md)|_Property_ > insideLeft|Represents the insideLeft value of plotArea.|beta|
|[chartPlotArea](../markdown/chartplotarea.md)|_Property_ > insideTop|Represents the insideTop value of plotArea.|beta|
|[chartPlotArea](../markdown/chartplotarea.md)|_Property_ > insideWidth|Represents the insideWidth value of plotArea.|beta|
|[chartPlotArea](../markdown/chartplotarea.md)|_Property_ > left|Represents the left value of plotArea.|beta|
|[chartPlotArea](../markdown/chartplotarea.md)|_Property_ > top|Represents the top value of plotArea.|beta|
|[chartPlotArea](../markdown/chartplotarea.md)|_Property_ > width|Represents the width value of plotArea.|beta|
|[chartPlotArea](../markdown/chartplotarea.md)|_Relationship_ > format|Represents the formatting of a chart plotArea. Read-only.|beta|
|[chartPlotArea](../markdown/chartplotarea.md)|_Relationship_ > position|Represents the position of plotArea.|beta|
|[chartPlotAreaFormat](../markdown/chartplotareaformat.md)|_Relationship_ > border|Represents the border attributes of a chart plotArea. Read-only.|beta|
|[chartPlotAreaFormat](../markdown/chartplotareaformat.md)|_Relationship_ > fill|Represents the fill format of an object, which includes background formating information. Read-only.|beta|
|[chartSeries](../markdown/chartseries.md)|_Property_ > axisGroup|Returns or sets the group for the specified series. ReadWrite Possible values are: Primary, Secondary.|beta|
|[chartSeries](../markdown/chartseries.md)|_Property_ > bubbleScale|Returns or sets the scale factor for bubbles in the specified chart group. Can be an integer value from 0 (zero) to 300, corresponding to a percentage of the default size. Applies only to bubble charts. ReadWrite.|beta|
|[chartSeries](../markdown/chartseries.md)|_Property_ > explosion|Returns or sets the explosion value for a pie-chart or doughnut-chart slice. Returns 0 (zero) if there's no explosion (the tip of the slice is in the center of the pie). ReadWrite.|beta|
|[chartSeries](../markdown/chartseries.md)|_Property_ > firstSliceAngle|Returns or sets the angle of the first pie-chart or doughnut-chart slice, in degrees (clockwise from vertical). Applies only to pie, 3-D pie, and doughnut charts. Can be a value from 0 through 360. ReadWrite|beta|
|[chartSeries](../markdown/chartseries.md)|_Property_ > hasLeaderLines|True if Microsoft Excel show leaderlines for each datalabel in series. ReadWrite.|beta|
|[chartSeries](../markdown/chartseries.md)|_Property_ > invertIfNegative|True if Microsoft Excel inverts the pattern in the item when it corresponds to a negative number. ReadWrite.|beta|
|[chartSeries](../markdown/chartseries.md)|_Property_ > overlap|Specifies how bars and columns are positioned. Can be a value between -100 and 100. Applies only to 2-D bar and 2-D column charts. ReadWrite.|beta|
|[chartSeries](../markdown/chartseries.md)|_Property_ > secondPlotSize|Returns or sets the size of the secondary section of either a pie of pie chart or a bar of pie chart, as a percentage of the size of the primary pie. Can be a value from 5 to 200. ReadWrite.|beta|
|[chartSeries](../markdown/chartseries.md)|_Property_ > splitType|Returns or sets the way the two sections of either a pie of pie chart or a bar of pie chart are split. ReadWrite. Possible values are: SplitByPosition, SplitByValue, SplitByPercentValue, SplitByCustomSplit.|beta|
|[chartSeries](../markdown/chartseries.md)|_Property_ > splitValue|Returns or sets the threshold value separating the two sections of either a pie of pie chart or a bar of pie chart. ReadWrite.|beta|
|[chartSeries](../markdown/chartseries.md)|_Property_ > type|Returns the value that represents the series type. Read-only. Possible values are: Column, Bar, Bar3D, Line, Pie, XYScatter, Area, Area3D, Doughnut, Radar, Surface3D, Column3D.|beta|
|[chartSeries](../markdown/chartseries.md)|_Property_ > varyByCategories|True if Microsoft Excel assigns a different color or pattern to each data marker. The chart must contain only one series. ReadWrite.|beta|
|[chartSeries](../markdown/chartseries.md)|_Relationship_ > dataLabels|Represents a collection of all dataLabels in the series. Read-only.|beta|
|[chartSeries](../markdown/chartseries.md)|_Relationship_ > xErrorBars|Represents the error bar object for a chart series. Read-only.|beta|
|[chartSeries](../markdown/chartseries.md)|_Relationship_ > yErrorBars|Represents the error bar object for a chart series. Read-only.|beta|
|[chartTrendline](../markdown/charttrendline.md)|_Property_ > backwardPeriod|Represents the number of periods that the trendline extends backward.|beta|
|[chartTrendline](../markdown/charttrendline.md)|_Property_ > forwardPeriod|Represents the number of periods that the trendline extends forward.|beta|
|[chartTrendline](../markdown/charttrendline.md)|_Property_ > showEquation|True if the equation for the trendline is displayed on the chart.|beta|
|[chartTrendline](../markdown/charttrendline.md)|_Property_ > showRSquared|True if the R-squared for the trendline is displayed on the chart.|beta|
|[chartTrendline](../markdown/charttrendline.md)|_Relationship_ > label|Represents the label of a chart trendline. Read-only.|beta|
|[chartTrendlineLabel](../markdown/charttrendlinelabel.md)|_Property_ > autoText|Boolean value representing if trendline label automatically generates appropriate text based on context.|beta|
|[chartTrendlineLabel](../markdown/charttrendlinelabel.md)|_Property_ > formula|String value that represents the formula of chart trendline label using A1-style notation.|beta|
|[chartTrendlineLabel](../markdown/charttrendlinelabel.md)|_Property_ > height|Returns the height, in points, of the chart trendline label. Read-only. Null if chart trendline label is not visible. Read-only.|beta|
|[chartTrendlineLabel](../markdown/charttrendlinelabel.md)|_Property_ > horizontalAlignment|Represents the horizontal alignment for chart trendline label. Possible values are: Center, Left, Right, Justify, Distributed.|beta|
|[chartTrendlineLabel](../markdown/charttrendlinelabel.md)|_Property_ > left|Represents the distance, in points, from the left edge of chart trendline label to the left edge of chart area. Null if chart trendline label is not visible.|beta|
|[chartTrendlineLabel](../markdown/charttrendlinelabel.md)|_Property_ > numberFormat|String value that represents the format code for trendline label.|beta|
|[chartTrendlineLabel](../markdown/charttrendlinelabel.md)|_Property_ > numberFormatLinked|Boolean value representing if the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells).|beta|
|[chartTrendlineLabel](../markdown/charttrendlinelabel.md)|_Property_ > text|String representing the text of the trendline label on a chart.|beta|
|[chartTrendlineLabel](../markdown/charttrendlinelabel.md)|_Property_ > textOrientation|Represents the text orientation of chart trendline label. The value should be an integer either from -90 to 90, or 180 for vertically-oriented text.|beta|
|[chartTrendlineLabel](../markdown/charttrendlinelabel.md)|_Property_ > top|Represents the distance, in points, from the top edge of chart trendline label to the top of chart area. Null if chart trendline label is not visible.|beta|
|[chartTrendlineLabel](../markdown/charttrendlinelabel.md)|_Property_ > verticalAlignment|Represents the vertical alignment of chart trendline label. Possible values are: Center, Bottom, Top, Justify, Distributed.|beta|
|[chartTrendlineLabel](../markdown/charttrendlinelabel.md)|_Property_ > width|Returns the width, in points, of the chart trendline label. Read-only. Null if chart trendline label is not visible. Read-only.|beta|
|[chartTrendlineLabel](../markdown/charttrendlinelabel.md)|_Relationship_ > format|Represents the format of chart trendline label. Read-only.|beta|
|[chartTrendlineLabelFormat](../markdown/charttrendlinelabelformat.md)|_Relationship_ > border|Represents the border format, which includes color, linestyle and weight. Read-only.|beta|
|[chartTrendlineLabelFormat](../markdown/charttrendlinelabelformat.md)|_Relationship_ > fill|Represents the fill format of the current chart trendline label. Read-only.|beta|
|[chartTrendlineLabelFormat](../markdown/charttrendlinelabelformat.md)|_Relationship_ > font|Represents the font attributes (font name, font size, color, etc.) for a chart trendline label. Read-only.|beta|
