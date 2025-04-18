import FormulaError from '../error';
import {H, Types} from '../helpers';

/**
 * Get the number of values that evaluate to true and false.
 * Cast Number and "TRUE", "FALSE" to boolean.
 * Ignore unrelated values.
 */
function getNumLogicalValue(params: any[]): [number, number] {
  let numTrue = 0, numFalse = 0;

  H.flattenParams(params, null, true, (val: any) => {
    const type = typeof val;
    let boolVal: boolean | null = null;

    if (type === "string") {
      if (val === 'TRUE') boolVal = true;
      else if (val === 'FALSE') boolVal = false;
    } else if (type === "number") {
      boolVal = Boolean(val);
    } else if (type === "boolean") {
      boolVal = val;
    }

    if (boolVal !== null) {
      boolVal ? numTrue++ : numFalse++;
    }
  });

  return [numTrue, numFalse];
}

// Interface for Logical Functions
interface LogicalFunctionsInterface {
  AND: (...params: any[]) => boolean;
  FALSE: () => boolean;
  IF: (context: any, logicalTest: any, valueIfTrue: any, valueIfFalse?: any) => any;
  IFERROR: (value: any, valueIfError: any) => any;
  IFNA: (value: any, valueIfNa: any) => any;
  IFS: (...params: any[]) => any;
  NOT: (logical: any) => boolean;
  OR: (...params: any[]) => boolean;
  TRUE: () => boolean;
  XOR: (...params: any[]) => boolean | FormulaError;
}

/**
 * Logical functions module with TypeScript type definitions
 */
const LogicalFunctions: LogicalFunctionsInterface = {
  /**
   * Returns TRUE if all arguments are TRUE
   */
  AND: (...params): boolean => {
    const [numTrue, numFalse] = getNumLogicalValue(params);

    // OR returns #VALUE! if no logical values are found.
    if (numTrue === 0 && numFalse === 0)
      throw FormulaError.VALUE;

    return numTrue > 0 && numFalse === 0;
  },

  /**
   * Returns the logical value FALSE
   */
  FALSE: (): boolean => false,

  /**
   * Tests a condition and returns one value if TRUE, another if FALSE
   */
  IF: (context, logicalTest, valueIfTrue, valueIfFalse = false): any => {
    logicalTest = H.accept(logicalTest, Types.BOOLEAN);
    valueIfTrue = H.accept(valueIfTrue); // do not parse type
    valueIfFalse = H.accept(valueIfFalse, null, false); // do not parse type

    return logicalTest ? valueIfTrue : valueIfFalse;
  },

  /**
   * Returns a specified value if the expression evaluates to an error; otherwise returns the result of the expression
   */
  IFERROR: (value, valueIfError): any => {
    return value.value instanceof FormulaError ? H.accept(valueIfError) : H.accept(value);
  },

  /**
   * Returns the specified value if the expression evaluates to #N/A, otherwise returns the result of the expression
   */
  IFNA: function (value, valueIfNa): any {
    if (arguments.length > 2)
      throw FormulaError.TOO_MANY_ARGS('IFNA');
    return FormulaError.NA.equals(value.value) ? H.accept(valueIfNa) : H.accept(value);
  },

  /**
   * Checks multiple conditions and returns a value corresponding to the first TRUE condition
   */
  IFS: (...params): any => {
    if (params.length % 2 !== 0)
      return new FormulaError('#N/A', 'IFS expects all arguments after position 0 to be in pairs.');

    for (let i = 0; i < params.length / 2; i++) {
      const logicalTest = H.accept(params[i * 2], Types.BOOLEAN);
      const valueIfTrue = H.accept(params[i * 2 + 1]);
      if (logicalTest)
        return valueIfTrue;
    }

    return FormulaError.NA;
  },

  /**
   * Reverses the logic of its argument
   */
  NOT: (logical): boolean => {
    logical = H.accept(logical, Types.BOOLEAN);
    return !logical;
  },

  /**
   * Returns TRUE if any argument is TRUE
   */
  OR: (...params): boolean => {
    const [numTrue, numFalse] = getNumLogicalValue(params);

    // OR returns #VALUE! if no logical values are found.
    if (numTrue === 0 && numFalse === 0)
      throw FormulaError.VALUE;

    return numTrue > 0;
  },

  /**
   * Returns the logical value TRUE
   */
  TRUE: (): boolean => true,

  /**
   * Returns a logical exclusive OR of all arguments
   */
  XOR: (...params): boolean | FormulaError => {
    const [numTrue, numFalse] = getNumLogicalValue(params);

    // XOR returns #VALUE! if no logical values are found.
    if (numTrue === 0 && numFalse === 0)
      return FormulaError.VALUE;

    return numTrue % 2 === 1;
  },
};

export default LogicalFunctions;
