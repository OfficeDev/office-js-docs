## Excel 1.3 Release Features

### Pivot Table resource and refresh action

The pivot table collection is now available on the worksheet resource. To begin with, this will enable developers to know about the pivot tables are available on a given workbook. This will also enable applications to refresh an individual pivot table or all pivot tables associated with a worksheet from the source data. Applications can take advantage of this feature by refreshing the pivot table remotely based on timing or other business criteria.

* Get pivotTable: Read properties of pivotTable object; currently limited to `name`. 
* `refresh()`: Refreshes a given pivotTable. Example, GET /worksheets/{id or name}/pivotTables/{id}/refresh
* `refreshAll()`: Refresh all tables within a given worksheet. Note that this action is available only on the pivot table collection. 
* In the future, we'll look to add more pivot table functionality extending the same objects.

### Visible rows on a filtered range

We've added a new resource called `visibleView` that represents visible rows of a given `range`. When a filter is applied on a range, it is often useful to know what values are visible (or selected from filter criteria). Previously, this required going through the underlying range and checking `hidden` property to determine whether a row is visible or not. Such a procedure is cumbersome when the range is fairly large.

With the new feature being added, applications can now apply a filter on a large table based on desired criteria and get easy access to filtered data. This feature will reduce the complexity involved in determining visible rows and take advantage of Excel's filtering capability.

`range.visibleView`

The visible view resource comes with useful properties that represent the range such as cell addresses, column and row count, index, formulas, number format, values/text, and value types.

If the visible range happens to be large, getting the entire resource might not be performant. In such a case, iterating over the rows provides a better user experience. The `rows` navigation property on visibleView allows just that - iterating over the visible range rows.


### New range manipulation functions

The following new range functions have been added. These functions reduce the complexity involved in determining the target range that the applications wishes to operate on.

Gets a certain number of rows or columns based on relative position of a given range:

* `columnsAfter(count: int)` 
* `columnsBefore(count: int)`
* `rowsAbove(count: int)`
* `rowsBelow(count: int)` 

`resizedRange`: Gets a range object similar to the current range object, but with its bottom-right corner expanded or contracted by some number of rows and columns.

### New table resource properties

Gain insights into the structure of the Excel table by using the following new properties: 

* `highlightFirstColumn`
* `highlightLastColumn`
* `showBandedColumns`
* `showBandedRows`
* `showFilterButton`

### Add binding

[`binding`](https://github.com/OfficeDev/office-js-docs/blob/master/reference/excel/bindingcollection.md) object represents an office.js binding that is defined in the workbook. It can be used to create a binding on a range that can later be used for tracking the range and also to regreister for events. 

Previously, we only allowed binding read methods. Now, we offer: 
* `add()` Add a new binding based on a named item in the workbook.
* `addFromNamedItem()` Add a new binding based on a named item in the workbook.
* `addFromSelection()` Add a new binding based on the current selection.

### Event management APIs in Excel namespace

Earlier, the event management APIs were only available through the common API namespace. Now, we've added the same API capability to register and manage *existing* events through the excel specific name space. This doesn't offer net new functionality - however this make easy to remain in the promise based batched API syntax reduce the dependency on common API for Excel related tasks. 

Currently, we offer `SelectionChanged`, `DataChanged`, `SettingsChanged` events. More to follow...

## Excel 1.4 Release Features

### Named item add and new properties

New properties
* `comment`
* `scope` worksheet or workbook scoped items 
* `worksheet` returns the worksheet on which the named item is scoped to.

New Methods
* `add(name: string, reference: Range or string, comment: string)`Adds a new name to the collection of the given scope.
* `addFormulaLocal(name: string, formula: string, comment: string)` Adds a new name to the collection of the given scope using the user's locale for the formula.

### Settings API in in Excel namespace

[Setting](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_1.4_OpenSpec/reference/excel/setting.md) object represents a key-value pair of a setting persisted to the document. Now, we've added settings related APIs under Excel namespace. This doesn't offer net new functionality - however this make easy to remain in the promise based batched API syntax reduce the dependency on common API for Excel related tasks. 

APIs include `getItem()` to get setting entry via the key, `add()` to add the specified key:value setting pair to the workbook.

### Others 

* Set table column name (prior version only allows reading).
* Add table column to the end of the table (prior version only allows anywhere but last).
* Add multiple rows to a table at a time (prior version only allows 1 row at a time). 
* `range.getColumnsAfter(count: number)` and `range.getColumnsBefore(count: number)` to get a certain number of columns to the right/left of the current Range object.
* Get item or null object function: This functionality allows getting object using a key. If the object does not exist, the returned object's isNullObject property will be true. This alows developers to check if an object exists or not without having to handle it thorugh exception handling. Available on worksheet, named-item, binding, chart series, etc.

`worksheet.GetItemOrNullObject()`

### Suspend calculation 
Suspends calculation until the next "context.sync()" is called. Once set, it is the developer's responsibility to re-calc the workbook, to ensure that any dependencies are propagated.

In addition, we are fixing the F9 re-calc bug that wasn't re-calculating the dirty cells.

## Code snippets

### Refresh pivot table

```js
  // Refresh pivot table
    async function queueRefreshPivotTable(sheet: Excel.Worksheet) {

        // Queue a command to get the table
        const expensesPivot = sheet.pivotTables.getItem("PivotTable3");

        // Make a change in the data
        const expensesTable = sheet.tables.getItem("ExpensesTable");

        let range = expensesTable.columns.getItem("AMOUNT").getDataBodyRange().getRow(2).values = [["$500"]];

        await sheet.context.sync();

        // Queue a command to refresh the pivot table 
        expensesPivot.refresh();

    }
```    

### List all named ranges

```js
/** List all named ranges in the workbook */
function listNamedRanges() {
  Excel.run(async (context) => {
    let sheet: Excel.Worksheet;
    try {

      // Load the property values of named ranges defined in the workbook
      const namedRanges = context.workbook.names.load();

       // Run the queued up commands
      await context.sync();

      // Console log all the list
      for (let i = 0; i < namedRanges.items.length; i++) {
        console.log(JSON.stringify(namedRanges.items[i]))+"\n";
      }

      //Queue a command to activate the sheet
      sheet.activate();

      // Run all of the queued up commands
      await context.sync();
    }
    catch (error) {
      OfficeHelpers.Utilities.log(error);
    }
  })
}
```

### Filter table data

```js
// Filter the table data
async function queueFilterData(sheet: Excel.Worksheet) {

        // Queue a command to get the table
        const expensesTable = sheet.tables.getItem("ExpensesTable");

        // Queue a command to filter and show only below average transactions 
            var filter = expensesTable.columns.getItem("AMOUNT").filter;
            filter.apply({
                filterOn: Excel.FilterOn.dynamic,
                dynamicCriteria: Excel.DynamicFilterCriteria.belowAverage
            });

}
```

### Sort table data

```js
    // Helper to sort the table data
    async function queueSortData(sheet: Excel.Worksheet) {

        // Queue a command to get the table
        const expensesTable = sheet.tables.getItem("ExpensesTable");

        // Queue a command to sort by most expensive transactions at the top (Amount, descending order)
            var sortRange = expensesTable.getDataBodyRange();
            sortRange.sort.apply([
				{
				    key: 3,
				    ascending: false,
				},
            ]);
    }
```

### Add a named range

```js
// Add a named range
async function queueAddName(sheet: Excel.Worksheet) {

        // Queue a command to add a table to the sheet
        let expensesTable = sheet.tables.add('A1:D1', true);
        expensesTable.name = "ExpensesTable";

        // Queue a command to set the header row
        expensesTable.getHeaderRowRange().values = [["DATE", "MERCHANT", "CATEGORY", "AMOUNT"]];

        let newData = transactions.map(item =>
          [item.DATE, item.MERCHANT, item.CATEGORY, item.AMOUNT]);

        expensesTable.rows.add(null, newData);

        // Add a named range
        sheet.names.add("TotalAmount", "=SUM(ExpensesTable[AMOUNT])");

        // Use the named range
        sheet.getRange("D13").values = [["=TotalAmount"]];

        //Autofit columns and rows
        sheet.getUsedRange().getEntireColumn().format.autofitColumns();
        sheet.getUsedRange().getEntireRow().format.autofitRows();
}
```


### Get visible range of a filtered table

```js
// Filter and get visible range of a filtered table
async function queueGetVisibleRange(sheet: Excel.Worksheet) {

        // Queue a command to get the table
         const expensesTable = sheet.tables.getItem("ExpensesTable");

        // Queue a command to filter and show only below average transactions 
           var filter = expensesTable.columns.getItem("AMOUNT").filter;

            filter.apply({
                filterOn: Excel.FilterOn.dynamic,
                dynamicCriteria: Excel.DynamicFilterCriteria.belowAverage
            });

            // Queue a command to get the visible view
            var visibleRange = expensesTable.getDataBodyRange().getVisibleView().load("values");

            await sheet.context.sync();
           
            const visibleValues = visibleRange.values;

            // console log the visible values
            console.log(visibleValues);
}
```