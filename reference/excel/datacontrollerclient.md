# DataControllerClient Object (JavaScript API for Excel)

Represents how the Visual is setup to use the data source.

## Properties

None

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[addField(wellId: number, fieldId: number, position: number)](#addfieldwellid-number-fieldid-number-position-number)|void|Add a field to well.|[ApiSet.InProgressFeatures.BiplatVisual](../requirement-sets/excel-api-requirement-sets.md)|
|[getAssociatedFields(wellId: number)](#getassociatedfieldswellid-number)|string|Gets an array of JSON objects representing the fields associated with the specified wellId.|[ApiSet.InProgressFeatures.BiplatVisual](../requirement-sets/excel-api-requirement-sets.md)|
|[getAvailableFields(wellId: number)](#getavailablefieldswellid-number)|string|Gets an array of JSON objects representing the fields that may be associated with wellId.|[ApiSet.InProgressFeatures.BiplatVisual](../requirement-sets/excel-api-requirement-sets.md)|
|[getWells()](#getwells)|string|Gets an array of JSON objects representing this visual's wells.|[ApiSet.InProgressFeatures.BiplatVisual](../requirement-sets/excel-api-requirement-sets.md)|
|[moveField(wellId: number, fromPosition: number, toPosition: number)](#movefieldwellid-number-fromposition-number-toposition-number)|void|Move a field from one position to another in a well.|[ApiSet.InProgressFeatures.BiplatVisual](../requirement-sets/excel-api-requirement-sets.md)|
|[removeField(wellId: number, position: number)](#removefieldwellid-number-position-number)|void|Remove a field from a well.|[ApiSet.InProgressFeatures.BiplatVisual](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### addField(wellId: number, fieldId: number, position: number)
Add a field to well.

#### Syntax
```js
dataControllerClientObject.addField(wellId, fieldId, position);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|wellId|number|The id of the well that the field will be added to.|
|fieldId|number|The id of the field to add to the well.|
|position|number|The position in the well where the field should be added.|

#### Returns
void

### getAssociatedFields(wellId: number)
Gets an array of JSON objects representing the fields associated with the specified wellId.

#### Syntax
```js
dataControllerClientObject.getAssociatedFields(wellId);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|wellId|number|The id of the well to get the associated fields for.|

#### Returns
string

### getAvailableFields(wellId: number)
Gets an array of JSON objects representing the fields that may be associated with wellId.

#### Syntax
```js
dataControllerClientObject.getAvailableFields(wellId);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|wellId|number|The id of the well to get the available fields for.|

#### Returns
string

### getWells()
Gets an array of JSON objects representing this visual's wells.

#### Syntax
```js
dataControllerClientObject.getWells();
```

#### Parameters
None

#### Returns
string

### moveField(wellId: number, fromPosition: number, toPosition: number)
Move a field from one position to another in a well.

#### Syntax
```js
dataControllerClientObject.moveField(wellId, fromPosition, toPosition);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|wellId|number|The id of the well to be moved.|
|fromPosition|number|The position in the well of the field to be moved.|
|toPosition|number|The new position for the field|

#### Returns
void

### removeField(wellId: number, position: number)
Remove a field from a well.

#### Syntax
```js
dataControllerClientObject.removeField(wellId, position);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|wellId|number|The id of the well that will have a field removed.|
|position|number|The position in the well of the field that should be removed|

#### Returns
void
