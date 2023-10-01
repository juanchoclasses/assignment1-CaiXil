import Cell from "./Cell"
import SheetMemory from "./SheetMemory"
import { ErrorMessages } from "./GlobalDefinitions";



export class FormulaEvaluator {
  // Define a function called update that takes a string parameter and returns a number
  private _errorOccured: boolean = false;
  private _errorMessage: string = "";
  private _currentFormula: FormulaType = [];
  private _lastResult: number = 0;
  private _sheetMemory: SheetMemory;
  private _result: number = 0;


  constructor(memory: SheetMemory) {
    this._sheetMemory = memory;
  }

  /**
   * 
   * @returns the result of the formula
   * 
   * 
   */

  private update(): number {
    // if there is error return the last result
    if (this._errorOccured) {
      return this._lastResult;
    }

    let result = this.extractTerm();

    while (this._currentFormula.length && (this._currentFormula[0] === "+" || this._currentFormula[0] === "-")) {
      let operator = this._currentFormula.shift();
      let term = this.extractTerm();
      if (operator === "-") {
        result -= term;
      } else {
        result += term;
      }
    }
    this._lastResult = result;
    return result;
  }

  private extractTerm(): number {
    // if there is error return the last result
    if (this._errorOccured) {
      return this._lastResult;
    }

    let result = this.extractFactor();

    while (this._currentFormula.length > 0 && (this._currentFormula[0] === "*" || this._currentFormula[0] === "/")) {
      let operator = this._currentFormula.shift();
      let factor = this.extractFactor();
      if (operator === "*") {
        result *= factor;
      } else if(operator === "/") {
        if (!factor) {
          this._errorOccured = true;
          this._errorMessage = ErrorMessages.divideByZero;
          this._lastResult = Infinity;
          return this._lastResult;
        } else {
          result /= factor;
        }
      }
    }
    // set the lastResult to the result
    this._lastResult = result;
    return result;
  }

  private extractFactor(): number {
    if (this._errorOccured) {
      return this._lastResult;
    }
    let result = 0;
    // if the formula is empty then set errorOccured to true, set the errorMessage to "PARTIAL" and return 0
    if (!this._currentFormula.length) {
      this._errorOccured = true;
      this._errorMessage = ErrorMessages.partial;
      return result;
    }

    // get the first token in the formula
    let token = this._currentFormula.shift();

    // if the token is a number set the result to the number, and set the lastResult to the number
    if (this.isNumber(token)) {
      result = Number(token);
      this._lastResult = result;

      // if the token is a left parenthesis, we can treat the content in the parentheses as a new formula
    } else if (token === "(") {
      result = this.update();
      if (!this._currentFormula.length || this._currentFormula.shift() !== ")") {
        this._errorOccured = true;
        this._errorMessage = ErrorMessages.missingParentheses;
        this._lastResult = result;
      }

      // if the token is a cell reference get the value of the cell
    } else if (this.isCellReference(token)) {
      [result, this._errorMessage] = this.getCellValue(token);

      // if the cell value is a number set the result to the number
      if (this._errorMessage) {
        this._errorOccured = true;
        this._lastResult = result;
      }

      // otherwise set the errorOccured flag to true
    } else {
      this._errorOccured = true;
      this._errorMessage = ErrorMessages.invalidFormula;
    }
    return result;
  }

  /**
    * place holder for the evaluator.   I am not sure what the type of the formula is yet 
    * I do know that there will be a list of tokens so i will return the length of the array
    * 
    * I also need to test the error display in the front end so i will set the error message to
    * the error messages found In GlobalDefinitions.ts
    * 
    * according to this formula.
    * 
    7 tokens partial: "#ERR",
    8 tokens divideByZero: "#DIV/0!",
    9 tokens invalidCell: "#REF!",
  10 tokens invalidFormula: "#ERR",
  11 tokens invalidNumber: "#ERR",
  12 tokens invalidOperator: "#ERR",
  13 missingParentheses: "#ERR",
  0 tokens emptyFormula: "#EMPTY!",

                    When i get back from my quest to save the world from the evil thing i will fix.
                      (if you are in a hurry you can fix it yourself)
                               Sincerely 
                               Bilbo
    * 
   */

  evaluate(formula: FormulaType) {

    // save the formula for later use
    this._currentFormula = [...formula];

    // return if the formula is empty
    if (!formula.length) {
      this._result = 0;
      this._errorMessage = ErrorMessages.emptyFormula;
      return;
    }

    //set errorOccured to false and errorMessage to empty string
    this._errorOccured = false;
    this._errorMessage = "";

    // evaluate the formula
    let result = this.update();
    this._result = result;

    // if there are still tokens in the formula set the errorOccured flag
    // if an error has occured then we dont update the error message
    if (this._currentFormula.length > 0 && !this._errorOccured) {
      this._errorOccured = true;
      this._errorMessage = ErrorMessages.invalidFormula;
    }

    // if an error occured  and the message is PARTIAL return the last result
    if (this._errorOccured) {
      this._result = this._lastResult;
    }
  }

  public get error(): string {
    return this._errorMessage
  }

  public get result(): number {
    return this._result;
  }




  /**
   * 
   * @param token 
   * @returns true if the toke can be parsed to a number
   */
  isNumber(token: TokenType): boolean {
    return !isNaN(Number(token));
  }

  /**
   * 
   * @param token
   * @returns true if the token is a cell reference
   * 
   */
  isCellReference(token: TokenType): boolean {

    return Cell.isValidCellLabel(token);
  }

  /**
   * 
   * @param token
   * @returns [value, ""] if the cell formula is not empty and has no error
   * @returns [0, error] if the cell has an error
   * @returns [0, ErrorMessages.invalidCell] if the cell formula is empty
   * 
   */
  getCellValue(token: TokenType): [number, string] {

    let cell = this._sheetMemory.getCellByLabel(token);
    let formula = cell.getFormula();
    let error = cell.getError();

    // if the cell has an error return 0
    if (error && error !== ErrorMessages.emptyFormula) {
      return [0, error];
    }

    // if the cell formula is empty return 0
    if (!formula.length) {
      return [0, ErrorMessages.invalidCell];
    }


    let value = cell.getValue();
    return [value, ""];

  }


}

export default FormulaEvaluator;