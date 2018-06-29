# VisualProperty Object (JavaScript API for Excel)

This object represents the attributes for a property.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|hasDefault|bool|Returns true when a default value for this property exists Read-only.|[ApiSet.InProgressFeatures.BiplatVisual](../requirement-sets/excel-api-requirement-sets.md)|
|id|string|Returns the property Id. Read-only.|[ApiSet.InProgressFeatures.BiplatVisual](../requirement-sets/excel-api-requirement-sets.md)|
|isDefault|bool|Returns true when the property's value is currently the default Read-only.|[ApiSet.InProgressFeatures.BiplatVisual](../requirement-sets/excel-api-requirement-sets.md)|
|localizedName|string|Returns the property localized name. Read-only.|[ApiSet.InProgressFeatures.BiplatVisual](../requirement-sets/excel-api-requirement-sets.md)|
|localizedOptions|string|Returns the localized property options for IEnumProperty only. If property type isn't enum, it returns null. Read-only.|[ApiSet.InProgressFeatures.BiplatVisual](../requirement-sets/excel-api-requirement-sets.md)|
|max|double|Returns max value of the property. Only valid for INumericProperty properties. Returns null if it's invalid. Read-only.|[ApiSet.InProgressFeatures.BiplatVisual](../requirement-sets/excel-api-requirement-sets.md)|
|min|double|Returns min value of the property. Only valid for INumericProperty properties. Returns null if it's invalid. Read-only.|[ApiSet.InProgressFeatures.BiplatVisual](../requirement-sets/excel-api-requirement-sets.md)|
|options|string|Returns the property options for IEnumProperty only. If property type isn't enum, it returns null. Read-only.|[ApiSet.InProgressFeatures.BiplatVisual](../requirement-sets/excel-api-requirement-sets.md)|
|stepSize|double|Returns step size of the property. Only valid for INumericProperty properties. Returns null if it's invalid. Read-only.|[ApiSet.InProgressFeatures.BiplatVisual](../requirement-sets/excel-api-requirement-sets.md)|
|value|object|Returns the property value. Read-only.|[ApiSet.InProgressFeatures.BiplatVisual](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|type|[VisualPropertyType](visualpropertytype.md)|Returns the property type. Read-only.|[ApiSet.InProgressFeatures.BiplatVisual](../requirement-sets/excel-api-requirement-sets.md)|

## Methods
None

