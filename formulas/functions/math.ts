import {FormulaError} from '../error';
import {H, Types, Criteria} from '../helpers';
import {Infix} from '../operators';

// Types for function context
interface FormulaContext {
  retrieveRanges: (range: any, sumRange: any) => [any[][], any[][]];
  retrieveArg: (arg: any) => any;
}

// Interface for Math Functions
interface MathFunctionsInterface {
  ABS: (number: any) => number;
  BASE: (number: any, radix: any, minLength?: any) => string;
  CEILING: (number: any, significance: any) => number;
  'CEILING.MATH': (number: any, significance?: any, mode?: any) => number;
  'CEILING.PRECISE': (number: any, significance?: any) => number;
  DECIMAL: (text: any, radix: any) => number;
  DEGREES: (radians: any) => number;
  EXP: (number: any) => number;
  FLOOR: (number: any, significance: any) => number;
  'FLOOR.MATH': (number: any, significance?: any, mode?: any) => number;
  'FLOOR.PRECISE': (number: any, significance?: any) => number;
  GCD: (...params: any[]) => number;
  INT: (number: any) => number;
  'ISO.CEILING': (...params: any[]) => number;
  LCM: (...params: any[]) => number;
  LN: (number: any) => number;
  LOG: (number: any, base?: any) => number;
  LOG10: (number: any) => number;
  MOD: (number: any, divisor: any) => number;
  MROUND: (number: any, multiple: any) => number;
  ODD: (number: any) => number;
  PI: () => number;
  POWER: (number: any, power: any) => number;
  PRODUCT: (...numbers: any[]) => number;
  QUOTIENT: (numerator: any, denominator: any) => number;
  RAND: () => number;
  RANDBETWEEN: (bottom: any, top: any) => number;
  ROUND: (number: any, digits: any) => number;
  ROUNDDOWN: (number: any, digits: any) => number;
  ROUNDUP: (number: any, digits: any) => number;
  SIGN: (number: any) => number;
  SQRT: (number: any) => number;
  SUM: (...params: any[]) => number;
  SUMIF: (context: FormulaContext, range: any, criteria: any, sumRange: any) => number;
  SUMPRODUCT: (array1: any, ...arrays: any[]) => number;
  SUMSQ: (...params: any[]) => number;
  TRUNC: (number: any) => number;
}

// Optimized MathFunctions module with ES-next features and TypeScript
const MathFunctions: MathFunctionsInterface = {
  ABS: (number): number =>
    Math.abs(H.accept(number, Types.NUMBER)),

  BASE: (number, radix, minLength = 0): string => {
    number = H.accept(number, Types.NUMBER);
    radix = H.accept(radix, Types.NUMBER);
    minLength = H.accept(minLength, Types.NUMBER, 0);

    if (number < 0 || number >= 2 ** 53 || radix < 2 || radix > 36 || minLength < 0) {
      throw FormulaError.NUM;
    }

    const result: string = number.toString(radix).toUpperCase();
    return '0'.repeat(Math.max(minLength - result.length, 0)) + result;
  },

  CEILING: (number, significance): number => {
    number = H.accept(number, Types.NUMBER);
    significance = H.accept(significance, Types.NUMBER);

    if (significance === 0) return 0;
    if (number / significance % 1 === 0) return number;

    const absSignificance: number = Math.abs(significance);
    const times: number = Math.floor(Math.abs(number) / absSignificance);

    return number < 0
      ? (significance < 0 ? -absSignificance * (times + 1) : -absSignificance * times)
      : (times + 1) * absSignificance;
  },

  'CEILING.MATH': (number, significance, mode = 0): number => {
    number = H.accept(number, Types.NUMBER);
    significance = H.accept(significance, Types.NUMBER, number > 0 ? 1 : -1);
    mode = H.accept(mode, Types.NUMBER, 0);

    if (number >= 0) {
      return MathFunctions.CEILING(number, significance);
    }

    const offset: number = mode ? significance : 0;
    return MathFunctions.CEILING(number, significance) - offset;
  },

  'CEILING.PRECISE': (number, significance = 1): number => {
    number = H.accept(number, Types.NUMBER);
    significance = H.accept(significance, Types.NUMBER, 1);
    return MathFunctions.CEILING(number, Math.abs(significance));
  },

  DECIMAL: (text, radix): number => {
    text = H.accept(text, Types.STRING);
    radix = Math.trunc(H.accept(radix, Types.NUMBER));

    if (radix < 2 || radix > 36) throw FormulaError.NUM;

    const res: number = parseInt(text, radix);
    if (isNaN(res)) throw FormulaError.NUM;

    return res;
  },

  DEGREES: (radians): number =>
    H.accept(radians, Types.NUMBER) * (180 / Math.PI),

  EXP: (number): number =>
    Math.exp(H.accept(number, Types.NUMBER)),

  FLOOR: (number, significance): number => {
    number = H.accept(number, Types.NUMBER);
    significance = H.accept(significance, Types.NUMBER);

    if (significance === 0) return 0;
    if (number > 0 && significance < 0) throw FormulaError.NUM;
    if (number / significance % 1 === 0) return number;

    const absSignificance: number = Math.abs(significance);
    const times: number = Math.floor(Math.abs(number) / absSignificance);

    return number < 0
      ? (significance < 0 ? -absSignificance * times : -absSignificance * (times + 1))
      : times * absSignificance;
  },

  'FLOOR.MATH': (number, significance, mode = 0): number => {
    number = H.accept(number, Types.NUMBER);
    significance = H.accept(significance, Types.NUMBER, number > 0 ? 1 : -1);
    mode = H.accept(mode, Types.NUMBER, 0);

    if (mode === 0 || number >= 0) {
      return MathFunctions.FLOOR(number, Math.abs(significance));
    }

    return MathFunctions.FLOOR(number, significance) + significance;
  },

  'FLOOR.PRECISE': (number, significance = 1): number => {
    number = H.accept(number, Types.NUMBER);
    significance = H.accept(significance, Types.NUMBER, 1);
    return MathFunctions.FLOOR(number, Math.abs(significance));
  },

  GCD: (...params): number => {
    // Extract and validate numbers
    const numbers: number[] = [];
    H.flattenParams(params, null, false, (param: any) => {
      const num: number = typeof param === 'boolean' ? NaN : Number(param);
      if (isNaN(num)) throw FormulaError.VALUE;
      if (num < 0 || num > 9007199254740990) throw FormulaError.NUM;
      numbers.push(Math.trunc(num));
    }, 0);

    // Optimized recursive GCD using ES-next features
    const gcd = (a: number, b: number): number => b === 0 ? a : gcd(b, a % b);

    // Reduce array using gcd function
    return numbers.reduce((result, num) => gcd(result, num), Math.abs(numbers[0] || 0));
  },

  INT: (number): number =>
    Math.floor(H.accept(number, Types.NUMBER)),

  'ISO.CEILING': (...params): number =>
    MathFunctions['CEILING.PRECISE'](...params),

  LCM: (...params): number => {
    // Extract and validate numbers
    const numbers: number[] = [];
    H.flattenParams(params, null, false, (param: any) => {
      const num: number = typeof param === 'boolean' ? NaN : Number(param);
      if (isNaN(num)) throw FormulaError.VALUE;
      if (num < 0 || num > 9007199254740990) throw FormulaError.NUM;
      numbers.push(Math.trunc(num));
    }, 1);

    // Optimized recursive GCD using ES-next features
    const gcd = (a: number, b: number): number => b === 0 ? a : gcd(b, a % b);

    // Find LCM using the relationship: LCM(a,b) = (a*b)/GCD(a,b)
    return numbers.reduce((lcm, num) => {
      const currGcd: number = gcd(lcm, num);
      return Math.abs(lcm * num) / currGcd;
    }, Math.abs(numbers[0] || 1));
  },

  LN: (number): number =>
    Math.log(H.accept(number, Types.NUMBER)),

  LOG: (number, base = 10): number => {
    number = H.accept(number, Types.NUMBER);
    base = H.accept(base, Types.NUMBER, 10);
    return Math.log(number) / Math.log(base);
  },

  LOG10: (number): number =>
    Math.log10(H.accept(number, Types.NUMBER)),

  MOD: (number, divisor): number => {
    number = H.accept(number, Types.NUMBER);
    divisor = H.accept(divisor, Types.NUMBER);
    if (divisor === 0) throw FormulaError.DIV0;
    return number - divisor * MathFunctions.INT(number / divisor);
  },

  MROUND: (number, multiple): number => {
    number = H.accept(number, Types.NUMBER);
    multiple = H.accept(multiple, Types.NUMBER);

    if (multiple === 0) return 0;
    if ((number > 0 && multiple < 0) || (number < 0 && multiple > 0)) throw FormulaError.NUM;
    if (number / multiple % 1 === 0) return number;

    return Math.round(number / multiple) * multiple;
  },

  ODD: (number): number => {
    number = H.accept(number, Types.NUMBER);
    if (number === 0) return 1;

    let temp: number = Math.ceil(Math.abs(number));
    temp = (temp & 1) ? temp : temp + 1;

    return (number > 0) ? temp : -temp;
  },

  PI: (): number => Math.PI,

  POWER: (number, power): number => {
    number = H.accept(number, Types.NUMBER);
    power = H.accept(power, Types.NUMBER);
    return number ** power;
  },

  PRODUCT: (...numbers): number => {
    let product: number = 1;
    H.flattenParams(numbers, null, true, (number: any, info: any) => {
      const parsedNumber: number = Number(number);
      if ((info.isLiteral && !isNaN(parsedNumber)) || typeof number === "number") {
        product *= info.isLiteral ? parsedNumber : number;
      }
    }, 1);
    return product;
  },

  QUOTIENT: (numerator, denominator): number => {
    numerator = H.accept(numerator, Types.NUMBER);
    denominator = H.accept(denominator, Types.NUMBER);
    return Math.trunc(numerator / denominator);
  },

  RAND: (): number => Math.random(),

  RANDBETWEEN: (bottom, top): number => {
    bottom = H.accept(bottom, Types.NUMBER);
    top = H.accept(top, Types.NUMBER);
    return Math.floor(Math.random() * (top - bottom + 1) + bottom);
  },

  ROUND: (number, digits): number => {
    number = H.accept(number, Types.NUMBER);
    digits = H.accept(digits, Types.NUMBER);

    const multiplier: number = 10 ** Math.abs(digits);
    const sign: number = number > 0 ? 1 : -1;

    if (digits > 0) {
      return sign * Math.round(Math.abs(number) * multiplier) / multiplier;
    } else if (digits === 0) {
      return sign * Math.round(Math.abs(number));
    } else {
      return sign * Math.round(Math.abs(number) / multiplier) * multiplier;
    }
  },

  ROUNDDOWN: (number, digits): number => {
    number = H.accept(number, Types.NUMBER);
    digits = H.accept(digits, Types.NUMBER);

    const multiplier: number = 10 ** Math.abs(digits);
    const sign: number = number > 0 ? 1 : -1;
    const offset: number = 0.5;

    if (digits > 0) {
      return sign * Math.round((Math.abs(number) - offset / multiplier) * multiplier) / multiplier;
    } else if (digits === 0) {
      return sign * Math.round(Math.abs(number) - offset);
    } else {
      return sign * Math.round((Math.abs(number) - offset * multiplier) / multiplier) * multiplier;
    }
  },

  ROUNDUP: (number, digits): number => {
    number = H.accept(number, Types.NUMBER);
    digits = H.accept(digits, Types.NUMBER);

    const multiplier: number = 10 ** Math.abs(digits);
    const sign: number = number > 0 ? 1 : -1;
    const offset: number = 0.5;

    if (digits > 0) {
      return sign * Math.round((Math.abs(number) + offset / multiplier) * multiplier) / multiplier;
    } else if (digits === 0) {
      return sign * Math.round(Math.abs(number) + offset);
    } else {
      return sign * Math.round((Math.abs(number) + offset * multiplier) / multiplier) * multiplier;
    }
  },

  SIGN: (number): number => {
    number = H.accept(number, Types.NUMBER);
    return number > 0 ? 1 : number === 0 ? 0 : -1;
  },

  SQRT: (number): number => {
    number = H.accept(number, Types.NUMBER);
    if (number < 0) throw FormulaError.NUM;
    return Math.sqrt(number);
  },

  SUM: (...params): number => {
    let result: number = 0;
    H.flattenParams(params, Types.NUMBER, true, (item: any, info: any) => {
      if (info.isLiteral || typeof item === "number") {
        result += info.isLiteral ? item : item;
      }
    });
    return result;
  },

  /**
   * This function requires instance of {@link FormulaParser}.
   */
  SUMIF: (context, range, criteria, sumRange): number => {
    const ranges = context.retrieveRanges(range, sumRange);
    range = ranges[0];
    sumRange = ranges[1];

    criteria = context.retrieveArg(criteria);
    const isCriteriaArray: boolean = criteria.isArray;
    // parse criteria
    criteria = Criteria.parse(H.accept(criteria));

    // Use array methods to calculate sum
    return range.reduce((sum: number, row: any[], rowNum: number) =>
        sum + row.reduce((rowSum: number, value: any, colNum: number) => {
          const valueToAdd = sumRange[rowNum][colNum];
          if (typeof valueToAdd !== "number") return rowSum;

          // Check if value matches criteria
          const matches: boolean = criteria.op === 'wc'
            ? criteria.match === criteria.value.test(value)
            : Infix.compareOp(value, criteria.op, criteria.value, Array.isArray(value), isCriteriaArray);

          return matches ? rowSum + valueToAdd : rowSum;
        }, 0),
      0);
  },

  SUMPRODUCT: (array1, ...arrays): number => {
    array1 = H.accept(array1, Types.ARRAY, undefined, false, true);

    // Process each additional array
    arrays.forEach(array => {
      array = H.accept(array, Types.ARRAY, undefined, false, true);

      // Verify dimensions match
      if (array1[0].length !== array[0].length || array1.length !== array.length)
        throw FormulaError.VALUE;

      // Multiply corresponding elements
      array1 = array1.map((row: any[], i: number) =>
        row.map((val: any, j: number) => {
          const a: number = typeof val === "number" ? val : 0;
          const b: number = typeof array[i][j] === "number" ? array[i][j] : 0;
          return a * b;
        })
      );
    });

    // Sum all elements using array methods
    return array1.reduce(
      (sum: number, row: any[]) => sum + row.reduce((rowSum: number, val: number) => rowSum + val, 0),
      0
    );
  },

  SUMSQ: (...params): number => {
    let result: number = 0;
    H.flattenParams(params, Types.NUMBER, true, (item: any, info: any) => {
      if (info.isLiteral || typeof item === "number") {
        result += (info.isLiteral ? item : item) ** 2;
      }
    });
    return result;
  },

  TRUNC: (number): number => Math.trunc(H.accept(number, Types.NUMBER)),
};

export default MathFunctions;
