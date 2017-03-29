```ts
namespace Excel {
  export class Range {
    dataValidation: DataValidation;
  }

  export class Worksheet {
    private _showInvalidDataValidation: boolean = false;
    get showInvalidDataValidation(): boolean {
        return this._showInvalidDataValidation;
    }
    set showInvalidDataValidation(show: boolean) {
        this._showInvalidDataValidation = show;
    }

    dataValidation: DataValidation;
  }

  // export interface RangeReference {
  //   address: string;
  // }

  error: DataValidationError;
  export class DataValidationError {
    constructor(
      public title?: string = 'The value you entered is not valid.'，
      public message?: string = 'A user has restricted values that can be entered into this cell.'，
      public style?: 'stop' | 'warning' | 'information' = 'stop'，
      public showOnInvalid?: boolean = true;
    ) { }
  }

  input: DataValidationInput;
  export class DataValidationInput {
    constructor(
      public title?: string = '',
      public message?: string = '',
      public showWhenSelected?: boolean = true
    ) { }
  }

  export interface ComparisonData {
    type: 'between' | 'notBetween' | 'equalTo' | 'notEqualTo' | 'greaterThan' | 'lessThan' | 'greaterThanOrEqualTo' | 'lessThanOrEqualTo',
    param1: number | string | Date | Range
    param2: number | string | Date | Range
  }

  export class DataValidation {
    status: 'empty' | 'same' | 'multiple';

    setRule(
      type: 'comparison' | 'list' | 'formula',
      data: string | [string | number | boolean] | Range | ComparisonData,
      ignoreBlank?: boolean = true
    ) {
      // setRule(type: 'comparison', data: {
      //   type: 'between' | 'notBetween' | 'equalTo' | 'notEqualTo' | 'greaterThan' | 'lessThan' | 'greaterThanOrEqualTo' | 'lessThanOrEqualTo',
      //   param1: number | string | Date | Range
      //   param2: number | string | Date | Range
      // });
      // setRule(type: 'list', data: string | [string | number | boolean] | Range);
      // setRule(type: 'formula', data: string);
    }

    getDataValidationOfSameType(): Range[] {
      return Range[];
    }

    getInvalidDataValidation(): Range[] {
      return Range[];
    }

    clearDataValidation(): void {

    }
  }
}


// Range demo
let range: Excel.Range;
range.dataValidation.setRule('comparison', { type: 'greaterThan', param1: 2 });

const comparisonData: ComparisonData = { type: 'between', param1: 2, param1: 5 };
range.dataValidation.setRule('comparison', comparisonData);

range.dataValidation.setRule('comparison', { ...comparisonData, param2: 10 });

if (range.status === 'empty' || 'same') {
  range.dataValidation.setRule('list', ['yes', 'no', 'maybe']);
}

range.dataValidation.error = new DataValidationError();
range.dataValidation.input = new DataValidationInput();

range.dataValidation.clearDataValidation();


// Worksheet demo
let worksheet: Excel.Worksheet;
worksheet.dataValidation.clearDataValidation();
worksheet.dataValidation.getDataValidation();
worksheet.dataValidation.getInvalidDataValidation();
```
