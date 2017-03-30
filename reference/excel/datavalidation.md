```ts
namespace Excel {
    export class Range {
        dataValidation: DataValidation;
    }
 
    export class DataValidation {
        type: { get: '' /*empty*/, 'comparison', 'list', 'formula', null /*inconsistent*/ }
 
        rule?: DataValidationRule
        input?: DataValidationInput;
        error?: DataValidationError;
        ignoreBlanks?: boolean;
 
        clear() { }
    }
 
    export interface DataValidationRule {
        comparison?: {
            type: 'double' | 'textLength' | 'date' | 'wholeNumber' /* | ... */,
            operator: 'between' | 'greaterThan' | 'equalTo' /* | ... */,
            formula1: number | string | Date | RangeReference | Range,
            formula2?: number | string | Date | RangeReference | Range
        },
        list?: {
            source: string | RangeReference | Range | [string | number | boolean],
            isDropdown?: boolean /* default = true */
        },
        custom?: {
            formula: string;
        }
    }
 
    export interface DataValidationError {
        style: 'stop' | 'warning' | 'information';
        title: string;
        message: string;
        show: boolean;
    }
 
    export interface DataValidationInput {
        show: boolean;
        title: string;
        message: string;
    }
 
    export interface RangeReference {
        address: string
    }
}
 
//////////////////////////
/////// EXAMPLES /////////
//////////////////////////
 
var range: Excel.Range;
 
// Set a new validation:
range.dataValidation.clear();
range.dataValidation.rule = {
    comparison: {
        type: 'wholeNumber',
        operator: "between",
        formula1: 5,
        formula2: 10
    }
};
range.dataValidation.error = {
    show: true,
    style: 'stop',
    title: "do it",
    message: "Please select something"
}
 
// On a different range:
range.dataValidation.rule = {
    list: {
        source: ['yes', 'no', 'maybe']
    }
}
```
