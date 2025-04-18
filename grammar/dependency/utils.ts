import {FormulaError} from '../../formulas/error';
import {Address} from '../../formulas/helpers';
import Collection from '../type/collection';
import {CellRef} from "../../types";

// Constants
const MAX_ROW = 1048576, MAX_COLUMN = 16384;

// Context interface for dependency parsing
export interface DepContext {
  position?: {
    sheet?: string;
    row?: number;
    col?: number;
  };
  retrieveRef: (obj: any) => any;
}

/**
 * Utility class for dependency parser
 */
export class Utils {
  context: DepContext;

  /**
   * Create a new Utils instance
   * @param context The dependency parser context
   */
  constructor(context: DepContext) {
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
  parseCellAddress(cellAddress: string): { ref: CellRef } {
    const res = cellAddress.match(/([$]?)([A-Za-z]{1,3})([$]?)([1-9][0-9]*)/);

    if (!res) {
      throw new Error(`Invalid cell address: ${cellAddress}`);
    }

    return {
      ref: {
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
  parseRow(row: string | number): { ref: CellRef } {
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
  parseCol(col: string): { ref: CellRef } {
    return {
      ref: {
        col: this.columnNameToNumber(col),
        row: undefined,
      },
    };
  }

  /**
   * Apply prefix operators to a value - for dependency tracking
   * @param prefixes Array of prefix operators
   * @param value Value to operate on
   * @return Constant value (not used for dependency tracking)
   */
  applyPrefix(prefixes: string[], value: any): number {
    this.extractRefValue(value);
    return 0;
  }

  /**
   * Apply postfix operators to a value - for dependency tracking
   * @param value Value to operate on
   * @param postfix Postfix operator
   * @return Constant value (not used for dependency tracking)
   */
  applyPostfix(value: any, postfix: string): number {
    this.extractRefValue(value);
    return 0;
  }

  /**
   * Apply infix operators to values - for dependency tracking
   * @param value1 First value
   * @param infix Infix operator
   * @param value2 Second value
   * @return Constant value (not used for dependency tracking)
   */
  applyInfix(value1: any, infix: string, value2: any): number {
    this.extractRefValue(value1);
    this.extractRefValue(value2);
    return 0;
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
    let maxRow: number, maxCol: number, minRow: number, minCol: number, sheet: string | undefined, res: { ref: any };

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
  applyUnion(refs: any[]): Collection {
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
  applyRange(refs: any[]): { ref: any } {
    let res: { ref: any }, maxRow = -1, maxCol = -1, minRow = MAX_ROW + 1, minCol = MAX_COLUMN + 1;

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
   * Extract value from a reference - for dependency tracking
   * @param obj Object to extract from
   * @returns Value and array flag
   */
  extractRefValue(obj: any): { val: any, isArray: boolean } {
    const isArray = Array.isArray(obj);

    if (obj && obj.ref) {
      // Can be number or array
      return {val: this.context.retrieveRef(obj), isArray};
    }

    return {val: obj, isArray};
  }

  /**
   * Convert to array
   * @param array Array to convert
   * @returns Original array
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
}

export default Utils;
