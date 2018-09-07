# PageBreakCollection Object (JavaScript API for Excel)

Represents the options in page layout margins.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|items|[PageBreak[]](pagebreak.md)|A collection of pageBreak objects. Read-only.|[beta](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[add(pageBreakRange: Range or string)](#addpagebreakrange-range-or-string)|[PageBreak](pagebreak.md)|Adds a page break before the top-left cell of the range specified.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[getCount()](#getcount)|int|Gets the number of page breaks in the collection.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[getItem(index: number)](#getitemindex-number)|[PageBreak](pagebreak.md)|Gets a page break object via the index.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|[removePageBreaks()](#removepagebreaks)|void|Resets all manual page breaks in the collection.|[beta](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### add(pageBreakRange: Range or string)
Adds a page break before the top-left cell of the range specified.

#### Syntax
```js
pageBreakCollectionObject.add(pageBreakRange);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|pageBreakRange|Range or string|The range immediately after the page break to be added.|

#### Returns
[PageBreak](pagebreak.md)

### getCount()
Gets the number of page breaks in the collection.

#### Syntax
```js
pageBreakCollectionObject.getCount();
```

#### Parameters
None

#### Returns
int

### getItem(index: number)
Gets a page break object via the index.

#### Syntax
```js
pageBreakCollectionObject.getItem(index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|index|number|Index of the page break.|

#### Returns
[PageBreak](pagebreak.md)

### removePageBreaks()
Resets all manual page breaks in the collection.

#### Syntax
```js
pageBreakCollectionObject.removePageBreaks();
```

#### Parameters
None

#### Returns
void
