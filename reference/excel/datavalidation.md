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
    status: 'none' | 'same' | 'multiple';

    constructor(
      public type: 'comparison' | 'list' | 'formula'，
      public param1: string | [string | number | boolean] | Range | ComparisonData,
      public param2?: string | [string | number | boolean] | Range | ComparisonData,
      public ignoreBlank?: boolean = true
    ) { }

    getDataValidationOfSameType(): Range[] {

    }

    getInvalidDataValidation(): Range[] {

    }

    clearDataValidation(): void {

    }
  }
}


// Range demo
let range: Excel.Range;

range.dataValidation = new DataValidation('comparison', { type: 'greaterThan', param1: 2 });

const comparisonData: ComparisonData = { type: 'between', param1: 2, param1: 5 };
range.dataValidation = new DataValidation('comparison', comparisonData);

range.dataValidation = new DataValidation('comparison', { ...comparisonData, param2: 10 });

if (range.status === 'none' || 'same') {
  range.dataValidation = new DataValidation('list', ['yes', 'no', 'maybe']);
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
