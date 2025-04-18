/**
 * Formula Error Class
 * Represents Excel-compatible formula errors with appropriate error codes
 */
class FormulaError extends Error {
  private _error: string;
  details?: object | Error;

  // Static error type map
  private static errorMap: Map<string, FormulaError> = new Map();

  /**
   * Creates a new formula error
   * @param error - Error code (e.g. "#DIV/0!")
   * @param msg - Optional detailed error message
   * @param details - Optional error details or original error object
   */
  constructor(error: string, msg?: string, details?: object | Error) {
    super(msg);

    // Return existing error instance if this is a standard error without details
    if (msg == null && details == null && FormulaError.errorMap.has(error)) {
      return FormulaError.errorMap.get(error) as FormulaError;
    }

    // Register new standard error in the map
    if (msg == null && details == null) {
      this._error = error;
      FormulaError.errorMap.set(error, this);
    } else {
      this._error = error;
    }

    this.details = details;
  }

  /**
   * Get the error name/code
   */
  get error(): string {
    return this._error;
  }

  /**
   * Get the error name/code (alias for error)
   */
  get name(): string {
    return this._error;
  }

  /**
   * Check if two errors are the same
   */
  equals(err: FormulaError): boolean {
    return err instanceof FormulaError && err._error === this._error;
  }

  /**
   * Convert to string representation
   */
  toString(): string {
    return this._error;
  }

  // Standard Excel formula errors
  static DIV0 = new FormulaError("#DIV/0!");
  static NA = new FormulaError("#N/A");
  static NAME = new FormulaError("#NAME?");
  static NULL = new FormulaError("#NULL!");
  static NUM = new FormulaError("#NUM!");
  static REF = new FormulaError("#REF!");
  static VALUE = new FormulaError("#VALUE!");

  /**
   * Create NOT_IMPLEMENTED error
   */
  static NOT_IMPLEMENTED = (functionName: string): FormulaError => {
    return new FormulaError("#NAME?", `Function ${functionName} is not implemented.`);
  };

  /**
   * Create TOO_MANY_ARGS error
   */
  static TOO_MANY_ARGS = (functionName: string): FormulaError => {
    return new FormulaError("#N/A", `Function ${functionName} has too many arguments.`);
  };

  /**
   * Create ARG_MISSING error
   */
  static ARG_MISSING = (args: number[]): FormulaError => {
    // Import here to avoid circular dependency
    // @ts-ignore
    const {Types} = require('./helpers');
    return new FormulaError("#N/A", `Argument type ${args.map(arg => Types[arg]).join(', ')} is missing.`);
  };

  /**
   * Create general ERROR
   */
  static ERROR = (msg: string, details?: object | Error): FormulaError => {
    return new FormulaError('#ERROR!', msg, details);
  };
}

export default FormulaError;
