import FormulaError from './error';
import {H} from './helpers';

/**
 * Value type mapping for comparison operations
 */
const type2Number: Record<string, number> = {
  'boolean': 3,
  'string': 2,
  'number': 1
};

/**
 * Prefix operators (unary + and -)
 */
export const Prefix = {
  /**
   * Apply unary operators (+ and -) to a value
   * @param prefixes Array of prefix operators
   * @param value The value to operate on
   * @param isArray Whether the value is an array
   */
  unaryOp: (prefixes: string[], value: any, isArray: boolean): any => {
    // Calculate final sign based on number of minus signs
    let sign = 1;
    for (const prefix of prefixes) {
      if (prefix === '-') {
        sign = -sign;
      } else if (prefix !== '+') {
        throw new Error(`Unrecognized prefix: ${prefix}`);
      }
    }

    // Default value is 0
    if (value == null) {
      value = 0;
    }

    // No change needed for positive sign
    if (sign === 1) {
      return value;
    }

    // Apply negative sign
    try {
      value = H.acceptNumber(value, isArray);
    } catch (e) {
      if (e instanceof FormulaError) {
        // Get first element for array if number parsing fails
        if (Array.isArray(value)) {
          value = value[0][0];
        }
      } else {
        throw e;
      }
    }

    // Check for NaN and return error if needed
    if (typeof value === "number" && isNaN(value)) {
      return FormulaError.VALUE;
    }

    return -value;
  }
};

/**
 * Postfix operators (% for percentage)
 */
export const Postfix = {
  /**
   * Apply postfix operator (%)
   * @param value The value to operate on
   * @param postfix The postfix operator
   * @param isArray Whether the value is an array
   */
  percentOp: (value: any, postfix: string, isArray: boolean): any => {
    try {
      value = H.acceptNumber(value, isArray);
    } catch (e) {
      if (e instanceof FormulaError) {
        return e;
      }
      throw e;
    }

    if (postfix === '%') {
      return value / 100;
    }

    throw new Error(`Unrecognized postfix: ${postfix}`);
  }
};

/**
 * Infix operators (comparison, concatenation, math)
 */
export const Infix = {
  /**
   * Apply comparison operators (=, <>, >, <, >=, <=)
   */
  compareOp: (value1: any, infix: string, value2: any, isArray1: boolean, isArray2: boolean): boolean => {
    // Set defaults and handle null values
    if (value1 == null) value1 = 0;
    if (value2 == null) value2 = 0;

    // Extract first element for arrays
    if (isArray1) {
      value1 = value1[0][0];
    }
    if (isArray2) {
      value2 = value2[0][0];
    }

    const type1 = typeof value1;
    const type2 = typeof value2;

    // Same type comparison
    if (type1 === type2) {
      switch (infix) {
        case '=':
          return value1 === value2;
        case '>':
          return value1 > value2;
        case '<':
          return value1 < value2;
        case '<>':
          return value1 !== value2;
        case '<=':
          return value1 <= value2;
        case '>=':
          return value1 >= value2;
      }
    }
    // Different type comparison - compare by type hierarchy
    else {
      switch (infix) {
        case '=':
          return false; // Different types are never equal
        case '>':
          return type2Number[type1] > type2Number[type2];
        case '<':
          return type2Number[type1] < type2Number[type2];
        case '<>':
          return true; // Different types are always not equal
        case '<=':
          return type2Number[type1] <= type2Number[type2];
        case '>=':
          return type2Number[type1] >= type2Number[type2];
      }
    }

    throw Error('Infix.compareOp: Should not reach here.');
  },

  /**
   * Apply concatenation operator (&)
   */
  concatOp: (value1: any, infix: string, value2: any, isArray1: boolean, isArray2: boolean): string => {
    // Set defaults and handle null values
    if (value1 == null) value1 = '';
    if (value2 == null) value2 = '';

    // Extract first element for arrays
    if (isArray1) {
      value1 = value1[0][0];
    }
    if (isArray2) {
      value2 = value2[0][0];
    }

    const type1 = typeof value1;
    const type2 = typeof value2;

    // Convert boolean to string representation
    if (type1 === 'boolean') {
      value1 = value1 ? 'TRUE' : 'FALSE';
    }
    if (type2 === 'boolean') {
      value2 = value2 ? 'TRUE' : 'FALSE';
    }

    // Concatenate as strings
    return '' + value1 + value2;
  },

  /**
   * Apply mathematical operators (+, -, *, /, ^)
   */
  mathOp: (value1: any, infix: string, value2: any, isArray1: boolean, isArray2: boolean): any => {
    // Set defaults and handle null values
    if (value1 == null) value1 = 0;
    if (value2 == null) value2 = 0;

    // Convert to numbers
    try {
      value1 = H.acceptNumber(value1, isArray1);
      value2 = H.acceptNumber(value2, isArray2);
    } catch (e) {
      if (e instanceof FormulaError) {
        return e;
      }
      throw e;
    }

    // Apply appropriate math operation
    switch (infix) {
      case '+':
        return value1 + value2;
      case '-':
        return value1 - value2;
      case '*':
        return value1 * value2;
      case '/':
        if (value2 === 0) {
          return FormulaError.DIV0;
        }
        return value1 / value2;
      case '^':
        return Math.pow(value1, value2);
    }

    throw Error('Infix.mathOp: Should not reach here.');
  },
};

/**
 * Exported operator categories
 */
export const Operators = {
  compareOp: ['<', '>', '=', '<>', '<=', '>='],
  concatOp: ['&'],
  mathOp: ['+', '-', '*', '/', '^'],
};

// Default exports
export default {
  Prefix,
  Postfix,
  Infix,
  Operators
};
