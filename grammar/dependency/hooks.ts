import {FormulaError} from '../../formulas/error';
import {H} from '../../formulas/helpers';
import {Parser} from '../parsing';
import Utils from './utils';
import GrammarUtils from '../utils';
import {lex} from "../lexing";

/**
 * Config for dependency parser
 */
export interface DepParserConfig {
  /**
   * Variable resolution callback
   * @param name Variable name
   * @param sheet Sheet name
   * @returns Variable reference
   */
  onVariable: (name: string, sheet?: string) => any;
}

/**
 * Dependency Parser class
 * Tracks cell and range dependencies in formulas
 */
export class DepParser {
  data: any[] = [];
  utils: Utils;
  onVariable: (name: string, sheet?: string) => any;
  functions: Record<string, any> = {};
  parser: Parser;
  position?: {
    sheet?: string;
    row?: number;
    col?: number;
  };

  /**
   * Create a new dependency parser
   * @param config Configuration options
   */
  constructor(config?: DepParserConfig) {
    this.data = [];
    config = {
      onVariable: () => null,
      ...config
    };

    this.utils = new Utils(this);
    this.onVariable = config.onVariable;
    this.functions = {};
    this.parser = new Parser(this, this.utils);
  }

  /**
   * Get value from a cell reference
   * @param ref Cell reference
   * @returns Placeholder value
   */
  getCell(ref: any): number {
    if (ref.row != null) {
      if (ref.sheet == null && this.position) {
        ref.sheet = this.position.sheet;
      }

      const idx = this.data.findIndex(element => {
        return (element.from && element.from.row <= ref.row && element.to.row >= ref.row
            && element.from.col <= ref.col && element.to.col >= ref.col)
          || (element.row === ref.row && element.col === ref.col && element.sheet === ref.sheet);
      });

      if (idx === -1) {
        this.data.push(ref);
      }
    }

    return 0;
  }

  /**
   * Get value from a range reference
   * @param ref Range reference
   * @returns Placeholder value
   */
  getRange(ref: any): number[][] {
    if (ref.from && ref.from.row != null) {
      if (ref.sheet == null && this.position) {
        ref.sheet = this.position.sheet;
      }

      const idx = this.data.findIndex(element => {
        return element.from && element.from.row === ref.from.row && element.from.col === ref.from.col
          && element.to.row === ref.to.row && element.to.col === ref.to.col;
      });

      if (idx === -1) {
        this.data.push(ref);
      }
    }

    return [[0]];
  }

  /**
   * Get variable reference
   * @param name Variable name
   * @returns Variable reference or error
   */
  getVariable(name: string): any {
    const res = {ref: this.onVariable(name, this.position?.sheet)};

    if (res.ref == null) {
      return FormulaError.NAME;
    }

    if (H.isCellRef(res)) {
      this.getCell(res.ref);
    } else {
      this.getRange(res.ref);
    }

    return 0;
  }

  /**
   * Retrieve references
   * @param valueOrRef Value or reference
   * @returns Value or reference result
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
   * Call a function (dummy implementation for dependency tracking)
   * @param name Function name
   * @param args Function arguments
   * @returns Placeholder result
   */
  callFunction(name: string, args: any[]): any {
    args.forEach(arg => {
      if (arg == null) {
        return;
      }
      this.retrieveRef(arg);
    });

    return {value: 0, ref: {}};
  }

  /**
   * Check formula result
   * @param result Formula result
   * @returns Result after checking
   */
  checkFormulaResult(result: any): any {
    this.retrieveRef(result);
  }

  /**
   * Parse a formula and extract dependencies
   * @param inputText Formula text
   * @param position Cell position (sheet, row, col)
   * @param ignoreError Whether to ignore errors
   * @returns Array of dependencies
   */
  parse(
    inputText: string,
    position: { sheet: string, row: number, col: number },
    ignoreError: boolean = false
  ): any[] {
    if (inputText.length === 0) {
      throw Error('Input must not be empty.');
    }

    this.data = [];
    this.position = position;

    const lexResult = lex(inputText);
    this.parser.input = lexResult.tokens;

    try {
      const res = this.parser.formulaWithBinaryOp();
      this.checkFormulaResult(res);
    } catch (e) {
      if (!ignoreError) {
        throw FormulaError.ERROR((e as Error).message, e);
      }
    }

    if (this.parser.errors.length > 0 && !ignoreError) {
      const error = this.parser.errors[0];
      throw GrammarUtils.formatChevrotainError(error, inputText);
    }

    return this.data;
  }
}

export default DepParser;
