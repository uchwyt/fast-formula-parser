import FormulaError from './error';
import Collection from "../grammar/type/collection";

/**
 * Data types used in formula parsing
 */
export enum Types {
  NUMBER = 0,
  ARRAY = 1,
  BOOLEAN = 2,
  STRING = 3,
  RANGE_REF = 4, // can be 'A:C' or '1:4', not only 'A1:C3'
  CELL_REF = 5,
  COLLECTIONS = 6, // Unions of references
  NUMBER_NO_BOOLEAN = 10,
}

// Types of parameters for formula processing
export interface ParamInfo {
  isLiteral: boolean;
  isCellRef: boolean;
  isRangeRef: boolean;
  isArray: boolean;
  isUnion: boolean;
}

// Function context interface
interface FormulaContext {
  utils: {
    extractRefValue: (obj: any) => { val: any, isArray: boolean };
  };
}

/**
 * Formula wildcard handling utilities
 */
export const WildCard = {
  /**
   * Check if a string contains wildcard characters
   */
  isWildCard: (obj: any): boolean => {
    if (typeof obj === "string")
      return /[*?]/.test(obj);
    return false;
  },

  /**
   * Convert a wildcard string to a regular expression
   */
  toRegex: (lookupText: string, flags?: string): RegExp => {
    return RegExp(lookupText.replace(/[.+^${}()|[\]\\]/g, '\\$&') // escape special chars for js regex
      .replace(/([^~]??)[?]/g, '$1.') // ? => .
      .replace(/([^~]??)[*]/g, '$1.*') // * => .*
      .replace(/~([?*])/g, '$1'), flags); // ~* => * and ~? => ?
  }
};

/**
 * Criteria parsing for filtering functions
 */
export const Criteria = {
  /**
   * Parse criteria string for comparison and wildcard matching
   */
  parse: (criteria: any): { op: string, value: any, match?: boolean } => {
    const type = typeof criteria;

    if (type === "string") {
      // Handle boolean strings
      const upper = criteria.toUpperCase();
      if (upper === 'TRUE' || upper === 'FALSE') {
        return {op: '=', value: upper === 'TRUE'};
      }

      // Handle comparison operators
      const res = criteria.match(/(<>|>=|<=|>|<|=)(.*)/);
      if (res) {
        const op = res[1];
        let value: any;

        // Handle different value types
        if (isNaN(Number(res[2]))) {
          const upper = res[2].toUpperCase();
          if (upper === 'TRUE' || upper === 'FALSE') {
            value = upper === 'TRUE';
          } else if (/#NULL!|#DIV\/0!|#VALUE!|#NAME\?|#NUM!|#N\/A|#REF!/.test(res[2])) {
            value = new FormulaError(res[2]);
          } else {
            value = res[2];
            if (WildCard.isWildCard(value)) {
              return {op: 'wc', value: WildCard.toRegex(value), match: op === '='};
            }
          }
        } else {
          value = Number(res[2]);
        }

        return {op, value};
      }

      // Handle plain wildcard
      if (WildCard.isWildCard(criteria)) {
        return {op: 'wc', value: WildCard.toRegex(criteria), match: true};
      }

      // Default to equality comparison
      return {op: '=', value: criteria};
    }

    // Handle non-string criteria
    if (type === "boolean" || type === 'number' || Array.isArray(criteria) || criteria instanceof FormulaError) {
      return {op: '=', value: criteria};
    }

    throw Error(`Criteria.parse: type ${typeof criteria} not supported`);
  }
};

/**
 * Address manipulation utilities
 */
export const Address = {
  /**
   * Convert column name to column number
   */
  columnNameToNumber: (columnName: string): number => {
    columnName = columnName.toUpperCase();
    const len = columnName.length;
    let number = 0;

    for (let i = 0; i < len; i++) {
      const code = columnName.charCodeAt(i);
      if (!isNaN(code)) {
        number += (code - 64) * Math.pow(26, len - i - 1);
      }
    }

    return number;
  },

  /**
   * Extend range2 to match dimensions in range1
   */
  extend: (range1: any, range2: any): any => {
    if (range2 == null) {
      return range1;
    }

    let rowOffset: number, colOffset: number;

    if (H.isCellRef(range1)) {
      rowOffset = 0;
      colOffset = 0;
    } else if (H.isRangeRef(range1)) {
      rowOffset = range1.ref.to.row - range1.ref.from.row;
      colOffset = range1.ref.to.col - range1.ref.from.col;
    } else {
      throw Error('Address.extend should not reach here.');
    }

    // Extend cell reference to match range dimensions
    if (H.isCellRef(range2)) {
      if (rowOffset > 0 || colOffset > 0) {
        range2 = {
          ref: {
            from: {col: range2.ref.col, row: range2.ref.row},
            to: {row: range2.ref.row + rowOffset, col: range2.ref.col + colOffset}
          }
        };
      }
    } else {
      // Extend range reference
      range2.ref.to.row = range2.ref.from.row + rowOffset;
      range2.ref.to.col = range2.ref.from.col + colOffset;
    }

    return range2;
  },
};

/**
 * Formula Helper methods
 */
class FormulaHelpers {
  Types = Types;
  type2Number: Record<string, number> = {
    number: Types.NUMBER,
    boolean: Types.BOOLEAN,
    string: Types.STRING,
    object: -1
  };

  /**
   * Check function result for errors
   */
  checkFunctionResult(result: any): any {
    const type = typeof result;

    if (type === 'number') {
      if (isNaN(result)) {
        return FormulaError.VALUE;
      } else if (!isFinite(result)) {
        return FormulaError.NUM;
      }
    }

    if (result === undefined || result === null) {
      return FormulaError.NULL;
    }

    return result;
  }

  /**
   * Deeply flatten an array
   */
  flattenDeep(arr: any[]): any[] {
    return arr.reduce((acc, val) =>
      Array.isArray(val) ? acc.concat(this.flattenDeep(val)) : acc.concat(val), []);
  }

  /**
   * Accept and normalize a number from various input types
   */
  acceptNumber(obj: any, isArray = true, allowBoolean = true) {
    // Check for error
    if (obj instanceof FormulaError) {
      return obj;
    }

    let number: number;

    if (typeof obj === 'number') {
      number = obj;
    }
    // TRUE -> 1, FALSE -> 0
    else if (typeof obj === 'boolean') {
      if (allowBoolean) {
        number = Number(obj);
      } else {
        throw FormulaError.VALUE;
      }
    }
    // "123" -> 123
    else if (typeof obj === 'string') {
      if (obj.length === 0) {
        throw FormulaError.VALUE;
      }
      number = Number(obj);
      // Check for NaN
      if (number !== number) {
        throw FormulaError.VALUE;
      }
    } else if (Array.isArray(obj)) {
      if (!isArray) {
        // For range ref, only allow single column range ref
        if (obj[0].length === 1) {
          number = this.acceptNumber(obj[0][0]) as number;
        } else {
          throw FormulaError.VALUE;
        }
      } else {
        number = this.acceptNumber(obj[0][0]) as number;
      }
    } else {
      throw Error('Unknown type in FormulaHelpers.acceptNumber');
    }

    return number;
  }

  /**
   * Flatten parameters and process with hook function
   */
  flattenParams(
    params: any[],
    valueType: Types | null,
    allowUnion: boolean,
    hook: (item: any, info: ParamInfo) => void,
    defValue: any = null,
    minSize = 1
  ): void {
    if (params.length < minSize) {
      throw FormulaError.ARG_MISSING([valueType as number]);
    }

    if (defValue == null) {
      defValue = valueType === Types.NUMBER ? 0 : valueType == null ? null : '';
    }

    params.forEach(param => {
      const {isCellRef, isRangeRef, isArray} = param;
      const isUnion = param.value instanceof Collection;
      const isLiteral = !isCellRef && !isRangeRef && !isArray && !isUnion;
      const info = {isLiteral, isCellRef, isRangeRef, isArray, isUnion};

      // Process different parameter types
      if (isLiteral) {
        if (param.omitted) {
          param = defValue;
        } else {
          param = this.accept(param, valueType, defValue);
        }
        hook(param, info);
      }
      // Cell reference
      else if (isCellRef) {
        hook(param.value, info);
      }
      // Union
      else if (isUnion) {
        if (!allowUnion) throw FormulaError.VALUE;
        param = param.value.data;
        param = this.flattenDeep(param);
        param.forEach(item => {
          hook(item, info);
        });
      }
      // Range reference or array
      else if (isRangeRef || isArray) {
        param = this.flattenDeep(param.value);
        param.forEach(item => {
          hook(item, info);
        });
      }
    });
  }

  /**
   * Accept and validate parameter with type checking
   */
  accept(
    param: any,
    type: Types | null = null,
    defValue?: any,
    flat = true,
    allowSingleValue = false
  ): any {
    // Handle array type
    if (Array.isArray(type)) {
      type = type[0];
    }

    // Handle missing parameter
    if (param == null && defValue === undefined) {
      throw FormulaError.ARG_MISSING([type as number]);
    } else if (param == null) {
      return defValue;
    }

    // Fast path for primitive types
    if (typeof param !== "object" || Array.isArray(param)) {
      return param;
    }

    const isArray = param.isArray;
    if (param.value != null) param = param.value;

    // Return unparsed type
    if (type == null) {
      return param;
    }

    // Handle formula errors
    if (param instanceof FormulaError) {
      throw param;
    }

    // Handle array type
    if (type === Types.ARRAY) {
      if (Array.isArray(param)) {
        return flat ? this.flattenDeep(param) : param;
      } else if (param instanceof Collection) {
        throw FormulaError.VALUE;
      } else if (allowSingleValue) {
        return flat ? [param] : [[param]];
      }
      throw FormulaError.VALUE;
    } else if (type === Types.COLLECTIONS) {
      return param;
    }

    // Extract first element from array
    if (isArray) {
      param = param[0][0];
    }

    // Type conversion based on expected type
    const paramType = this.type(param);
    if (type === Types.STRING) {
      if (paramType === Types.BOOLEAN) {
        param = param ? 'TRUE' : 'FALSE';
      } else {
        param = `${param}`;
      }
    } else if (type === Types.BOOLEAN) {
      if (paramType === Types.STRING) {
        throw FormulaError.VALUE;
      }
      if (paramType === Types.NUMBER) {
        param = Boolean(param);
      }
    } else if (type === Types.NUMBER) {
      param = this.acceptNumber(param, false);
    } else if (type === Types.NUMBER_NO_BOOLEAN) {
      param = this.acceptNumber(param, false, false);
    } else {
      throw FormulaError.VALUE;
    }

    return param;
  }

  /**
   * Determine the type of a value
   */
  type(variable: any): number {
    let type = this.type2Number[typeof variable];

    if (type === -1) {
      if (Array.isArray(variable)) {
        type = Types.ARRAY;
      } else if (variable.ref) {
        if (variable.ref.from) {
          type = Types.RANGE_REF;
        } else {
          type = Types.CELL_REF;
        }
      } else if (variable instanceof Collection) {
        type = Types.COLLECTIONS;
      }
    }

    return type;
  }

  /**
   * Check if param is a range reference
   */
  isRangeRef(param: any): boolean {
    return param.ref && param.ref.from;
  }

  /**
   * Check if param is a cell reference
   */
  isCellRef(param: any): boolean {
    return param.ref && !param.ref.from;
  }

  /**
   * Helper function for retrieving ranges
   */
  retrieveRanges(context: FormulaContext, range1: any, range2: any): [any, any] {
    // Process args
    range2 = Address.extend(range1, range2);

    // Retrieve values
    range1 = this.retrieveArg(context, range1);
    range1 = H.accept(range1, Types.ARRAY, undefined, false, true);

    if (range2 !== range1) {
      range2 = this.retrieveArg(context, range2);
      range2 = H.accept(range2, Types.ARRAY, undefined, false, true);
    } else {
      range2 = range1;
    }

    return [range1, range2];
  }

  /**
   * Retrieve argument value
   */
  retrieveArg(context: FormulaContext, arg: any): any {
    if (arg === null) {
      return {value: 0, isArray: false, omitted: true};
    }

    const res = context.utils.extractRefValue(arg);
    return {value: res.val, isArray: res.isArray, ref: arg.ref};
  }
}

// Create singleton instance
const H = new FormulaHelpers();

// Export classes and instances
export {FormulaHelpers, H};
