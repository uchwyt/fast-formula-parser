import {FormulaError} from '../formulas/error';
import {Address} from '../formulas/helpers';
import {Prefix, Postfix, Infix, Operators} from '../formulas/operators';
import Collection from './type/collection';
import {NotAllInputParsedException} from 'chevrotain';
import type {ParserContext, Reference} from "../types";

// Constants
const MAX_ROW = 1048576;
const MAX_COLUMN = 16384;


/**
 * Utils class for formula parsing
 */
export class Utils {
  context: ParserContext;

  /**
   * Create a new Utils instance
   * @param context The parser context
   */
  constructor(context: ParserContext) {
    this.context = context;
  }

  /**
   * Convert column name to number
   * @param columnName Column name (e.g., "A", "BC")
   * @returns Column number
   */
  columnNameToNumber(columnName: string): number {
    return Address.columnNameToNumber(columnName);
  }

  /**
   * Parse a cell address string
   * @param cellAddress Cell address string (e.g., "A1", "$B$2")
   * @returns Cell reference object
   */
  parseCellAddress(cellAddress: string): Reference {
    const res = cellAddress.match(/([$]?)([A-Za-z]{1,3})([$]?)([1-9][0-9]*)/);

    if (!res) {
      throw new Error(`Invalid cell address: ${cellAddress}`);
    }

    return {
      ref: {
        address: res[0],
        col: this.columnNameToNumber(res[2]),
        row: +res[4]
      },
    };
  }

  /**
   * Parse a row reference
   * @param row Row number
   * @returns Row reference object
   */
  parseRow(row: string | number): Reference {
    const rowNum = +row;
    if (!Number.isInteger(rowNum)) {
      throw Error('Row number must be integer.');
    }

    return {
      ref: {
        col: undefined,
        row: rowNum
      },
    };
  }

  /**
   * Parse a column reference
   * @param col Column name
   * @returns Column reference object
   */
  parseCol(col: string): Reference {
    return {
      ref: {
        col: this.columnNameToNumber(col),
        row: undefined,
      },
    };
  }

  /**
   * Parse a column range
   * @param col1 First column name
   * @param col2 Second column name
   * @returns Range reference object
   */
  parseColRange(col1: string, col2: string): Reference {
    const col1Num = this.columnNameToNumber(col1);
    const col2Num = this.columnNameToNumber(col2);

    return {
      ref: {
        from: {
          col: Math.min(col1Num, col2Num),
          row: null as any
        },
        to: {
          col: Math.max(col1Num, col2Num),
          row: null as any
        }
      }
    };
  }

  /**
   * Parse a row range
   * @param row1 First row number
   * @param row2 Second row number
   * @returns Range reference object
   */
  parseRowRange(row1: number, row2: number): Reference {
    return {
      ref: {
        from: {
          col: null as any,
          row: Math.min(row1, row2),
        },
        to: {
          col: null as any,
          row: Math.max(row1, row2),
        }
      }
    };
  }

  /**
   * Apply prefix operators to a value
   * @param prefixes Array of prefix operators
   * @param val Value to operate on
   * @param isArray Whether the value is an array
   * @returns Result after applying prefix operators
   */
  private _applyPrefix(prefixes: string[], val: any, isArray: boolean): any {
    if (this.isFormulaError(val)) {
      return val;
    }
    return Prefix.unaryOp(prefixes, val, isArray);
  }

  /**
   * Apply prefix operators (async version)
   */
  async applyPrefixAsync(prefixes: string[], value: any): Promise<any> {
    const {val, isArray} = this.extractRefValue(await value);
    return this._applyPrefix(prefixes, val, isArray);
  }

  /**
   * Apply prefix operators
   */
  applyPrefix(prefixes: string[], value: any): any {
    if (this.context.async) {
      return this.applyPrefixAsync(prefixes, value);
    } else {
      const {val, isArray} = this.extractRefValue(value);
      return this._applyPrefix(prefixes, val, isArray);
    }
  }

  /**
   * Apply postfix operators to a value
   * @param val Value to operate on
   * @param isArray Whether the value is an array
   * @param postfix Postfix operator
   * @returns Result after applying postfix operator
   */
  private _applyPostfix(val: any, isArray: boolean, postfix: string): any {
    if (this.isFormulaError(val)) {
      return val;
    }
    return Postfix.percentOp(val, postfix, isArray);
  }

  /**
   * Apply postfix operators (async version)
   */
  async applyPostfixAsync(value: any, postfix: string): Promise<any> {
    const {val, isArray} = this.extractRefValue(await value);
    return this._applyPostfix(val, isArray, postfix);
  }

  /**
   * Apply postfix operators
   */
  applyPostfix(value: any, postfix: string): any {
    if (this.context.async) {
      return this.applyPostfixAsync(value, postfix);
    } else {
      const {val, isArray} = this.extractRefValue(value);
      return this._applyPostfix(val, isArray, postfix);
    }
  }

  /**
   * Apply infix operators to values
   * @param res1 First value result object
   * @param infix Infix operator
   * @param res2 Second value result object
   * @returns Result after applying infix operator
   */
  private _applyInfix(res1: { val: any, isArray: boolean }, infix: string, res2: { val: any, isArray: boolean }): any {
    const val1 = res1.val, isArray1 = res1.isArray;
    const val2 = res2.val, isArray2 = res2.isArray;

    if (this.isFormulaError(val1)) {
      return val1;
    }
    if (this.isFormulaError(val2)) {
      return val2;
    }

    if (Operators.compareOp.includes(infix)) {
      return Infix.compareOp(val1, infix, val2, isArray1, isArray2);
    } else if (Operators.concatOp.includes(infix)) {
      return Infix.concatOp(val1, infix, val2, isArray1, isArray2);
    } else if (Operators.mathOp.includes(infix)) {
      return Infix.mathOp(val1, infix, val2, isArray1, isArray2);
    } else {
      throw new Error(`Unrecognized infix: ${infix}`);
    }
  }

  /**
   * Apply infix operators (async version)
   */
  async applyInfixAsync(value1: any, infix: string, value2: any): Promise<any> {
    const res1 = this.extractRefValue(await value1);
    const res2 = this.extractRefValue(await value2);
    return this._applyInfix(res1, infix, res2);
  }

  /**
   * Apply infix operators
   */
  applyInfix(value1: any, infix: string, value2: any): any {
    if (this.context.async) {
      return this.applyInfixAsync(value1, infix, value2);
    } else {
      const res1 = this.extractRefValue(value1);
      const res2 = this.extractRefValue(value2);
      return this._applyInfix(res1, infix, res2);
    }
  }

  /**
   * Apply intersection operator to references
   * @param refs Array of references to intersect
   * @returns Intersection result
   */
  applyIntersect(refs: any[]): any {
    if (this.isFormulaError(refs[0])) {
      return refs[0];
    }

    if (!refs[0].ref) {
      throw Error(`Expecting a reference, but got ${refs[0]}.`);
    }

    // A intersection will keep track of references, value won't be retrieved here
    let maxRow: number, maxCol: number, minRow: number, minCol: number, sheet: string | undefined, res: Reference;

    // First time setup
    const ref = refs.shift().ref;
    sheet = ref.sheet;

    if (!ref.from) {
      // Check whole row/col reference
      if (ref.row === undefined || ref.col === undefined) {
        throw Error('Cannot intersect the whole row or column.');
      }

      // Cell ref
      maxRow = minRow = ref.row;
      maxCol = minCol = ref.col;
    } else {
      // Range ref
      maxRow = Math.max(ref.from.row, ref.to.row);
      minRow = Math.min(ref.from.row, ref.to.row);
      maxCol = Math.max(ref.from.col, ref.to.col);
      minCol = Math.min(ref.from.col, ref.to.col);
    }

    let err;
    refs.forEach(ref => {
      if (this.isFormulaError(ref)) {
        return ref;
      }

      ref = ref.ref;
      if (!ref) throw Error(`Expecting a reference, but got ${ref}.`);

      if (!ref.from) {
        if (ref.row === undefined || ref.col === undefined) {
          throw Error('Cannot intersect the whole row or column.');
        }

        // Cell ref
        if (ref.row > maxRow || ref.row < minRow || ref.col > maxCol || ref.col < minCol
          || sheet !== ref.sheet) {
          err = FormulaError.NULL;
        }

        maxRow = minRow = ref.row;
        maxCol = minCol = ref.col;
      } else {
        // Range ref
        const refMaxRow = Math.max(ref.from.row, ref.to.row);
        const refMinRow = Math.min(ref.from.row, ref.to.row);
        const refMaxCol = Math.max(ref.from.col, ref.to.col);
        const refMinCol = Math.min(ref.from.col, ref.to.col);

        if (refMinRow > maxRow || refMaxRow < minRow || refMinCol > maxCol || refMaxCol < minCol
          || sheet !== ref.sheet) {
          err = FormulaError.NULL;
        }

        // Update
        maxRow = Math.min(maxRow, refMaxRow);
        minRow = Math.max(minRow, refMinRow);
        maxCol = Math.min(maxCol, refMaxCol);
        minCol = Math.max(minCol, refMinCol);
      }
    });

    if (err) return err;

    // Check if the ref can be reduced to cell reference
    if (maxRow === minRow && maxCol === minCol) {
      res = {
        ref: {
          sheet,
          row: maxRow,
          col: maxCol
        }
      };
    } else {
      res = {
        ref: {
          sheet,
          from: {row: minRow, col: minCol},
          to: {row: maxRow, col: maxCol}
        }
      };
    }

    if (!res.ref.sheet) {
      delete res.ref.sheet;
    }

    return res;
  }

  /**
   * Apply union operator to references
   * @param refs Array of references to union
   * @returns Union collection
   */
  applyUnion(refs: any[]): Collection | FormulaError {
    const collection = new Collection();

    for (let i = 0; i < refs.length; i++) {
      if (this.isFormulaError(refs[i])) {
        return refs[i];
      }
      collection.add(this.extractRefValue(refs[i]).val, refs[i]);
    }

    return collection;
  }

  /**
   * Apply range operator to references
   * @param refs Array of references
   * @returns Range reference
   */
  applyRange(refs: any[]): Reference {
    let res: Reference, maxRow = -1, maxCol = -1, minRow = MAX_ROW + 1, minCol = MAX_COLUMN + 1;

    refs.forEach(ref => {
      if (this.isFormulaError(ref)) {
        return ref;
      }

      // Row ref is saved as number, parse the number to row ref here
      if (typeof ref === 'number') {
        ref = this.parseRow(ref);
      }

      ref = ref.ref;

      // Check whole row/col reference
      if (ref.row === undefined) {
        minRow = 1;
        maxRow = MAX_ROW;
      }

      if (ref.col === undefined) {
        minCol = 1;
        maxCol = MAX_COLUMN;
      }

      if (ref.row > maxRow) {
        maxRow = ref.row;
      }

      if (ref.row < minRow) {
        minRow = ref.row;
      }

      if (ref.col > maxCol) {
        maxCol = ref.col;
      }

      if (ref.col < minCol) {
        minCol = ref.col;
      }
    });

    if (maxRow === minRow && maxCol === minCol) {
      res = {
        ref: {
          row: maxRow,
          col: maxCol
        }
      };
    } else {
      res = {
        ref: {
          from: {row: minRow, col: minCol},
          to: {row: maxRow, col: maxCol}
        }
      };
    }

    return res;
  }

  /**
   * Extract value from a reference
   * @param obj Object to extract from
   * @returns Value and array flag
   */
  extractRefValue(obj: any): { val: any, isArray: boolean } {
    let res = obj;
    let isArray = false;

    if (Array.isArray(res)) {
      isArray = true;
    }

    if (obj && obj.ref) {
      // Can be number or array
      return {val: this.context.retrieveRef(obj), isArray};
    }

    return {val: res, isArray};
  }

  /**
   * Convert to array
   * @param array Array to convert
   * @returns Converted array
   */
  toArray(array: any[]): any[] {
    return array;
  }

  /**
   * Convert to number
   * @param number Number string
   * @returns Number
   */
  toNumber(number: string): number {
    return Number(number);
  }

  /**
   * Convert to string
   * @param string String with quotes
   * @returns String without quotes
   */
  toString(string: string): string {
    return string.substring(1, string.length - 1).replace(/""/g, '"');
  }

  /**
   * Convert to boolean
   * @param bool Boolean string
   * @returns Boolean value
   */
  toBoolean(bool: string): boolean {
    return bool.toUpperCase() === 'TRUE';
  }

  /**
   * Convert to formula error
   * @param error Error string
   * @returns FormulaError
   */
  toError(error: string): FormulaError {
    return new FormulaError(error.toUpperCase());
  }

  /**
   * Check if object is a formula error
   * @param obj Object to check
   * @returns True if object is a formula error
   */
  isFormulaError(obj: any): boolean {
    return obj instanceof FormulaError;
  }

  /**
   * Format Chevrotain error with location information
   * @param error Chevrotain error
   * @param inputText Input text
   * @returns Formula error with location information
   */
  static formatChevrotainError(error: any, inputText: string): FormulaError {
    let line: number, column: number, msg = '';

    // E.g. SUM(1))
    if (error instanceof NotAllInputParsedException) {
      line = error.token.startLine;
      column = error.token.startColumn;
    } else {
      line = error.previousToken.startLine;
      column = error.previousToken.startColumn + 1;
    }

    msg += '\n' + inputText.split('\n')[line - 1] + '\n';
    msg += Array(column - 1).fill(' ').join('') + '^\n';
    msg += `Error at position ${line}:${column}\n` + error.message;
    error.errorLocation = {line, column};

    return FormulaError.ERROR(msg, error);
  }
}

export default Utils;
