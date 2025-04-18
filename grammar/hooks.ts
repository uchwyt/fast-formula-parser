import TextFunctions from '../formulas/functions/text';
import MathFunctions from '../formulas/functions/math';
import LogicalFunctions from '../formulas/functions/logical';
import EngFunctions from '../formulas/functions/engineering';
import InformationFunctions from '../formulas/functions/information';
import StatisticalFunctions from '../formulas/functions/statistical';
import DateFunctions from '../formulas/functions/date';
import {FormulaError} from '../formulas/error';
import {H} from '../formulas/helpers';
import {Parser} from './parsing';
import {allTokenNames, lex} from './lexing';
import Utils from './utils';

/**
 * Config for formula parser
 */
export interface FormulaParserConfig {
  /**
   * Custom functions
   */
  functions?: Record<string, (...args: any[]) => any>;

  /**
   * Functions that need context
   */
  functionsNeedContext?: Record<string, (context: any, ...args: any[]) => any>;

  /**
   * Variable resolution callback
   * @param name Variable name
   * @param sheet Sheet name
   * @param position Cell position
   */
  onVariable?: (name: string, sheet?: string, position?: any) => any;

  /**
   * Cell value provider
   * @param ref Cell reference
   */
  onCell?: (ref: any) => any;

  /**
   * Range value provider
   * @param ref Range reference
   */
  onRange?: (ref: any) => any[][];
}

/**
 * Formula Parser
 * Parses and evaluates Excel-like formulas
 */
export class FormulaParser {
  logs: string[] = [];
  isTest: boolean;
  utils: Utils;
  onVariable: (name: string, sheet?: string, position?: any) => any;
  functions: Record<string, (...args: any[]) => any>;
  onRange: (ref: any) => any[][];
  onCell: (ref: any) => any;
  funsNullAs0: string[];
  funsNeedContextAndNoDataRetrieve: string[];
  funsNeedContext: string[];
  funsPreserveRef: string[];
  parser: Parser;
  position?: {
    sheet?: string;
    row?: number;
    col?: number;
  };
  async?: boolean;

  /**
   * Create a new formula parser
   * @param config Parser configuration
   * @param isTest Whether in test mode
   */
  constructor(config?: FormulaParserConfig, isTest: boolean = false) {
    this.logs = [];
    this.isTest = isTest;
    this.utils = new Utils(this);

    const defaultConfig: FormulaParserConfig = {
      functions: {},
      functionsNeedContext: {},
      onVariable: () => null,
      onCell: () => 0,
      onRange: () => [[0]],
    };

    config = {...defaultConfig, ...config};

    this.onVariable = config.onVariable!;
    this.functions = {
      ...DateFunctions,
      ...StatisticalFunctions,
      ...InformationFunctions,
      ...EngFunctions,
      ...LogicalFunctions,
      ...TextFunctions,
      ...MathFunctions,
      ...config.functions,
      ...config.functionsNeedContext
    };

    this.onRange = config.onRange!;
    this.onCell = config.onCell!;

    // Functions treat null as 0, other functions treat null as ""
    this.funsNullAs0 = [
      ...Object.keys(MathFunctions),
      ...Object.keys(LogicalFunctions),
      ...Object.keys(EngFunctions),
      ...Object.keys(StatisticalFunctions),
      ...Object.keys(DateFunctions)
    ];

    // Functions need context and don't need to retrieve references
    this.funsNeedContextAndNoDataRetrieve = [
      'ROW', 'ROWS', 'COLUMN', 'COLUMNS', 'SUMIF', 'INDEX', 'AVERAGEIF', 'IF'
    ];

    // Functions need parser context
    this.funsNeedContext = [
      ...Object.keys(config.functionsNeedContext || {}),
      ...this.funsNeedContextAndNoDataRetrieve,
      'INDEX', 'OFFSET', 'INDIRECT', 'IF', 'CHOOSE', 'WEBSERVICE'
    ];

    // Functions preserve reference in arguments
    this.funsPreserveRef = Object.keys(InformationFunctions);

    this.parser = new Parser(this, this.utils);
  }

  /**
   * Get all lexing token names
   * @returns All token names
   */
  static get allTokens(): string[] {
    return allTokenNames;
  }

  /**
   * Get value from a cell reference
   * @param ref Cell reference
   * @returns Cell value
   */
  getCell(ref: any): any {
    if (ref.sheet == null && this.position) {
      ref.sheet = this.position.sheet;
    }

    return this.onCell(ref);
  }

  /**
   * Get values from a range reference
   * @param ref Range reference
   * @returns Range values
   */
  getRange(ref: any): any[][] {
    if (ref.sheet == null && this.position) {
      ref.sheet = this.position.sheet;
    }

    return this.onRange(ref);
  }

  /**
   * Get a variable value
   * @param name Variable name
   * @returns Variable value or reference
   */
  getVariable(name: string): any {
    const res = {
      ref: this.onVariable(name, this.position?.sheet, this.position)
    };

    if (res.ref == null) {
      return FormulaError.NAME;
    }

    return res;
  }

  /**
   * Retrieve value from a reference
   * @param valueOrRef Value or reference
   * @returns Value from reference
   */
  retrieveRef(valueOrRef: any): any {
    if (H.isRangeRef(valueOrRef)) {
      return this.getRange(valueOrRef.ref);
    }

    if (H.isCellRef(valueOrRef)) {
      return this.getCell(valueOrRef.ref);
    }

    return valueOrRef;
  }

  /**
   * Call a function
   * @param name Function name
   * @param args Function arguments
   * @returns Function result
   */
  private _callFunction(name: string, args: any[]): any {
    if (name.startsWith('_xlfn.')) {
      name = name.slice(6);
    }

    name = name.toUpperCase();

    // If one arg is null, it means 0 or "" depends on the function it calls
    const nullValue = this.funsNullAs0.includes(name) ? 0 : '';

    if (!this.funsNeedContextAndNoDataRetrieve.includes(name)) {
      // Retrieve reference
      args = args.map(arg => {
        if (arg === null) {
          return {value: nullValue, isArray: false, omitted: true};
        }

        const res = this.utils.extractRefValue(arg);

        if (this.funsPreserveRef.includes(name)) {
          return {value: res.val, isArray: res.isArray, ref: arg.ref};
        }

        return {
          value: res.val,
          isArray: res.isArray,
          isRangeRef: !!H.isRangeRef(arg),
          isCellRef: !!H.isCellRef(arg)
        };
      });
    }

    if (this.functions[name]) {
      let res;
      try {
        if (!this.funsNeedContextAndNoDataRetrieve.includes(name) && !this.funsNeedContext.includes(name)) {
          res = this.functions[name](...args);
        } else {
          res = this.functions[name](this, ...args);
        }
      } catch (e) {
        // Allow functions throw FormulaError, this makes functions easier to implement!
        if (e instanceof FormulaError) {
          return e;
        } else {
          throw e;
        }
      }

      if (res === undefined) {
        if (this.isTest) {
          if (!this.logs.includes(name)) this.logs.push(name);
          return {value: 0, ref: {}};
        }

        throw FormulaError.NOT_IMPLEMENTED(name);
      }

      return res;
    } else {
      if (this.isTest) {
        if (!this.logs.includes(name)) this.logs.push(name);
        return {value: 0, ref: {}};
      }

      throw FormulaError.NOT_IMPLEMENTED(name);
    }
  }

  /**
   * Call a function asynchronously
   * @param name Function name
   * @param args Function arguments
   * @returns Promise of function result
   */
  async callFunctionAsync(name: string, args: any[]): Promise<any> {
    const awaitedArgs = [];

    for (const arg of args) {
      awaitedArgs.push(await arg);
    }

    const res = await this._callFunction(name, awaitedArgs);
    return H.checkFunctionResult(res);
  }

  /**
   * Call a function
   * @param name Function name
   * @param args Function arguments
   * @returns Function result
   */
  callFunction(name: string, args: any[]): any {
    if (this.async) {
      return this.callFunctionAsync(name, args);
    } else {
      const res = this._callFunction(name, args);
      return H.checkFunctionResult(res);
    }
  }

  /**
   * Return currently supported functions
   * @returns Array of supported function names
   */
  supportedFunctions(): string[] {
    const supported: string[] = [];
    const functions = Object.keys(this.functions);

    functions.forEach(fun => {
      try {
        const res = this.functions[fun](0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);
        if (res === undefined) return;
        supported.push(fun);
      } catch (e) {
        if (e instanceof Error) {
          supported.push(fun);
        }
      }
    });

    return supported.sort();
  }

  /**
   * Check and return the appropriate formula result
   * @param result Formula result
   * @param allowReturnArray Whether to allow array results
   * @returns Processed formula result
   */
  checkFormulaResult(result: any, allowReturnArray: boolean = false): any {
    const type = typeof result;

    // Number
    if (type === 'number') {
      if (isNaN(result)) {
        return FormulaError.VALUE;
      } else if (!isFinite(result)) {
        return FormulaError.NUM;
      }
      result += 0; // make -0 to 0
    }
    // Object
    else if (type === 'object') {
      if (result instanceof FormulaError) {
        return result;
      }

      if (allowReturnArray) {
        if (result.ref) {
          result = this.retrieveRef(result);
        }

        // Disallow union, and other unknown data types.
        // e.g. `=(A1:C1, A2:E9)` -> #VALUE!
        if (typeof result === 'object' && !Array.isArray(result) && result != null) {
          return FormulaError.VALUE;
        }
      } else {
        if (result.ref && result.ref.row && !result.ref.from) {
          // Single cell reference
          result = this.retrieveRef(result);
        } else if (result.ref && result.ref.from && result.ref.from.col === result.ref.to.col) {
          // Single column reference
          result = this.retrieveRef({
            ref: {
              row: result.ref.from.row,
              col: result.ref.from.col
            }
          });
        } else if (Array.isArray(result)) {
          result = result[0][0];
        } else {
          // Array, range reference, union collections
          return FormulaError.VALUE;
        }
      }
    }

    return result;
  }

  /**
   * Parse an Excel formula
   * @param inputText Formula text
   * @param position Cell position
   * @param allowReturnArray Whether to allow array results
   * @returns Parsed result
   */
  parse(
    inputText: string,
    position?: { row: number, col: number, sheet?: string },
    allowReturnArray: boolean = false
  ): any {
    if (inputText.length === 0) {
      throw Error('Input must not be empty.');
    }

    this.position = position;
    this.async = false;

    const lexResult = lex(inputText);
    this.parser.input = lexResult.tokens;

    let res;
    try {
      res = this.parser.formulaWithBinaryOp();
      res = this.checkFormulaResult(res, allowReturnArray);

      if (res instanceof FormulaError) {
        return res;
      }
    } catch (e) {
      throw FormulaError.ERROR((e as Error).message, e);
    }

    if (this.parser.errors.length > 0) {
      const error = this.parser.errors[0];
      throw Utils.formatChevrotainError(error, inputText);
    }

    return res;
  }

  /**
   * Parse an Excel formula asynchronously
   * @param inputText Formula text
   * @param position Cell position
   * @param allowReturnArray Whether to allow array results
   * @returns Promise of parsed result
   */
  async parseAsync(
    inputText: string,
    position?: { row: number, col: number, sheet?: string },
    allowReturnArray: boolean = false
  ): Promise<any> {
    if (inputText.length === 0) {
      throw Error('Input must not be empty.');
    }

    this.position = position;
    this.async = true;

    const lexResult = lex(inputText);
    this.parser.input = lexResult.tokens;

    let res;
    try {
      res = await this.parser.formulaWithBinaryOp();
      res = this.checkFormulaResult(res, allowReturnArray);

      if (res instanceof FormulaError) {
        return res;
      }
    } catch (e) {
      throw FormulaError.ERROR((e as Error).message, e);
    }

    if (this.parser.errors.length > 0) {
      const error = this.parser.errors[0];
      throw Utils.formatChevrotainError(error, inputText);
    }

    return res;
  }
}

export default FormulaParser;
