import {tokenVocabulary} from './lexing';
import {EmbeddedActionsParser} from 'chevrotain';
import {Utils} from './utils';
import type {ParserContext} from "../types";

const {
  String,
  SheetQuoted,
  Function,
  FormulaErrorT,
  RefError,
  Cell,
  Sheet,
  Name,
  Number,
  Boolean,
  Column,

  Comma,
  Colon,
  Semicolon,
  OpenParen,
  CloseParen,
  OpenCurlyParen,
  CloseCurlyParen,
  MulOp,
  PlusOp,
  DivOp,
  MinOp,
  ConcatOp,
  ExOp,
  PercentOp,
  NeqOp,
  GteOp,
  LteOp,
  GtOp,
  EqOp,
  LtOp
} = tokenVocabulary;

/**
 * Parser class for formula parsing
 */
export class Parser extends EmbeddedActionsParser {
  context: ParserContext;
  utils: Utils;
  binaryOperatorsPrecedence: string[][];
  c1: any;

  /**
   * Create a new Parser instance
   * @param context Parser context
   * @param utils Utils instance
   */
  constructor(context: ParserContext, utils: Utils) {
    super(tokenVocabulary as any, {
//      outputCst: false,
      maxLookahead: 1,
      skipValidations: true,
    });

    this.context = context;
    this.utils = utils;
    this.binaryOperatorsPrecedence = [
      ['^'],
      ['*', '/'],
      ['+', '-'],
      ['&'],
      ['<', '>', '=', '<>', '<=', '>='],
    ];

    this.initializeRules();
  }

  /**
   * Get serialized grammar for visualization
   */
  getSerializedGastProductions(): any {
    return this.getGAstProductions();
  }

  /**
   * Initialize parser rules
   */
  private initializeRules(): void {
    const $ = this;

    // Formula with binary operators
    $.RULE('formulaWithBinaryOp', () => {
      const infixes: string[] = [];
      const values: any[] = [$.SUBRULE($.formulaWithPercentOp)];

      $.MANY(() => {
        // Caching Arrays of Alternatives
        // https://sap.github.io/chevrotain/docs/guide/performance.html#caching-arrays-of-alternatives
        infixes.push($.OR($.c1 ||
          (
            $.c1 = [
              {ALT: () => $.CONSUME(GtOp).image},
              {ALT: () => $.CONSUME(EqOp).image},
              {ALT: () => $.CONSUME(LtOp).image},
              {ALT: () => $.CONSUME(NeqOp).image},
              {ALT: () => $.CONSUME(GteOp).image},
              {ALT: () => $.CONSUME(LteOp).image},
              {ALT: () => $.CONSUME(ConcatOp).image},
              {ALT: () => $.CONSUME(PlusOp).image},
              {ALT: () => $.CONSUME(MinOp).image},
              {ALT: () => $.CONSUME(MulOp).image},
              {ALT: () => $.CONSUME(DivOp).image},
              {ALT: () => $.CONSUME(ExOp).image}
            ]
          )));
        values.push($.SUBRULE2($.formulaWithPercentOp));
      });

      $.ACTION(() => {
        // evaluate
        for (const ops of this.binaryOperatorsPrecedence) {
          for (let index = 0, length = infixes.length; index < length; index++) {
            const infix = infixes[index];
            if (!ops.includes(infix)) continue;

            infixes.splice(index, 1);
            values.splice(index, 2, this.utils.applyInfix(values[index], infix, values[index + 1]));
            index--;
            length--;
          }
        }
      });

      return values[0];
    });

    // Plus/minus operators
    $.RULE('plusMinusOp', () => $.OR([
      {ALT: () => $.CONSUME(PlusOp).image},
      {ALT: () => $.CONSUME(MinOp).image}
    ]));

    // Formula with percent operator
    $.RULE('formulaWithPercentOp', () => {
      let value = $.SUBRULE($.formulaWithUnaryOp);

      $.OPTION(() => {
        const postfix = $.CONSUME(PercentOp).image;
        value = $.ACTION(() => this.utils.applyPostfix(value, postfix));
      });

      return value;
    });

    // Formula with unary operators
    $.RULE('formulaWithUnaryOp', () => {
      // Support ++---3 => -3
      const prefixes: string[] = [];

      $.MANY(() => {
        const op = $.OR([
          {ALT: () => $.CONSUME(PlusOp).image},
          {ALT: () => $.CONSUME(MinOp).image}
        ]);

        prefixes.push(op);
      });

      const formula = $.SUBRULE($.formulaWithIntersect);

      if (prefixes.length > 0) {
        return $.ACTION(() => this.utils.applyPrefix(prefixes, formula));
      }

      return formula;
    });

    // Formula with intersection
    $.RULE('formulaWithIntersect', () => {
      // e.g.  'A1 A2 A3'
      let ref1 = $.SUBRULE($.formulaWithRange);
      const refs = [ref1];

      $.MANY({
        GATE: () => {
          // See https://github.com/SAP/chevrotain/blob/master/examples/grammars/css/css.js#L436-L441
          const prevToken = $.LA(0);
          const nextToken = $.LA(1);
          // This is the only place where the grammar is whitespace sensitive
          return nextToken.startOffset > prevToken.endOffset + 1;
        },
        DEF: () => {
          refs.push($.SUBRULE3($.formulaWithRange));
        }
      });

      if (refs.length > 1) {
        return $.ACTION(() => this.utils.applyIntersect(refs));
      }

      return ref1;
    });

    // Formula with range
    $.RULE('formulaWithRange', () => {
      // e.g. 'A1:C3' or 'A1:A3:C4', can be any number of references, at least 2
      const ref1 = $.SUBRULE($.formula);
      const refs = [ref1];

      $.MANY(() => {
        $.CONSUME(Colon);
        refs.push($.SUBRULE2($.formula));
      });

      if (refs.length > 1) {
        return $.ACTION(() => this.utils.applyRange(refs));
      }

      return ref1;
    });

    // Formula
    $.RULE('formula', () => $.OR9([
      {ALT: () => $.SUBRULE($.referenceWithoutInfix)},
      {ALT: () => $.SUBRULE($.paren)},
      {ALT: () => $.SUBRULE($.constant)},
      {ALT: () => $.SUBRULE($.functionCall)},
      {ALT: () => $.SUBRULE($.constantArray)},
    ]));

    // Parentheses
    $.RULE('paren', () => {
      // Formula paren or union paren
      $.CONSUME(OpenParen);
      let result;
      const refs = [];

      refs.push($.SUBRULE($.formulaWithBinaryOp));

      $.MANY(() => {
        $.CONSUME(Comma);
        refs.push($.SUBRULE2($.formulaWithBinaryOp));
      });

      if (refs.length > 1) {
        result = $.ACTION(() => this.utils.applyUnion(refs));
      } else {
        result = refs[0];
      }

      $.CONSUME(CloseParen);

      return result;
    });

    // Constant array
    $.RULE('constantArray', () => {
      const arr: any[][] = [[]];
      let currentRow = 0;

      $.CONSUME(OpenCurlyParen);

      // Array must contain at least one item
      arr[currentRow].push($.SUBRULE($.constantForArray));

      $.MANY(() => {
        const sep = $.OR([
          {ALT: () => $.CONSUME(Comma).image},
          {ALT: () => $.CONSUME(Semicolon).image}
        ]);

        const constant = $.SUBRULE2($.constantForArray);

        if (sep === ',') {
          arr[currentRow].push(constant);
        } else {
          currentRow++;
          arr[currentRow] = [];
          arr[currentRow].push(constant);
        }
      });

      $.CONSUME(CloseCurlyParen);

      return $.ACTION(() => this.utils.toArray(arr));
    });

    // Constant for array
    $.RULE('constantForArray', () => $.OR([
      {
        ALT: () => {
          const prefix = $.OPTION(() => $.SUBRULE($.plusMinusOp));
          const image = $.CONSUME(Number).image;
          const number = $.ACTION(() => this.utils.toNumber(image));

          if (prefix) {
            return $.ACTION(() => this.utils.applyPrefix([prefix], number));
          }

          return number;
        }
      }, {
        ALT: () => {
          const str = $.CONSUME(String).image;
          return $.ACTION(() => this.utils.toString(str));
        }
      }, {
        ALT: () => {
          const bool = $.CONSUME(Boolean).image;
          return $.ACTION(() => this.utils.toBoolean(bool));
        }
      }, {
        ALT: () => {
          const err = $.CONSUME(FormulaErrorT).image;
          return $.ACTION(() => this.utils.toError(err));
        }
      }, {
        ALT: () => {
          const err = $.CONSUME(RefError).image;
          return $.ACTION(() => this.utils.toError(err));
        }
      },
    ]));

    // Constant
    $.RULE('constant', () => $.OR([
      {
        ALT: () => {
          const number = $.CONSUME(Number).image;
          return $.ACTION(() => this.utils.toNumber(number));
        }
      }, {
        ALT: () => {
          const str = $.CONSUME(String).image;
          return $.ACTION(() => this.utils.toString(str));
        }
      }, {
        ALT: () => {
          const bool = $.CONSUME(Boolean).image;
          return $.ACTION(() => this.utils.toBoolean(bool));
        }
      }, {
        ALT: () => {
          const err = $.CONSUME(FormulaErrorT).image;
          return $.ACTION(() => this.utils.toError(err));
        }
      },
    ]));

    // Function call
    $.RULE('functionCall', () => {
      const functionName = $.CONSUME(Function).image.slice(0, -1);
      const args = $.SUBRULE($.arguments);

      $.CONSUME(CloseParen);

      // Dependency parser won't call function
      return $.ACTION(() => this.context.callFunction(functionName, args));
    });

    // Function arguments
    $.RULE('arguments', () => {
      // Allows ',' in the front
      $.MANY2(() => {
        $.CONSUME2(Comma);
      });

      const args: any[] = [];

      // Allows empty arguments
      $.OPTION(() => {
        args.push($.SUBRULE($.formulaWithBinaryOp));

        $.MANY(() => {
          $.CONSUME1(Comma);
          args.push(null); // e.g. ROUND(1.5,)

          $.OPTION3(() => {
            args.pop();
            args.push($.SUBRULE2($.formulaWithBinaryOp));
          });
        });
      });

      return args;
    });

    // Reference without infix
    $.RULE('referenceWithoutInfix', () => $.OR([
      {ALT: () => $.SUBRULE($.referenceItem)},
      {
        // Sheet name prefix
        ALT: () => {
          const sheetName = $.SUBRULE($.prefixName);
          const referenceItem = $.SUBRULE2($.formulaWithRange);

          $.ACTION(() => {
            if (this.utils.isFormulaError(referenceItem)) {
              return referenceItem;
            }

            referenceItem.ref.sheet = sheetName;
          });

          return referenceItem;
        }
      },
    ]));

    // Reference item
    $.RULE('referenceItem', () => $.OR([
      {
        ALT: () => {
          const address = $.CONSUME(Cell).image;
          return $.ACTION(() => this.utils.parseCellAddress(address));
        }
      },
      {
        ALT: () => {
          const name = $.CONSUME(Name).image;
          return $.ACTION(() => this.context.getVariable(name));
        }
      },
      {
        ALT: () => {
          const column = $.CONSUME(Column).image;
          return $.ACTION(() => this.utils.parseCol(column));
        }
      },
      // A row check should be here, but the token is same with Number,
      // In order to resolve ambiguities, I leave this empty, and
      // parse the number to row number when needed.
      {
        ALT: () => {
          const err = $.CONSUME(RefError).image;
          return $.ACTION(() => this.utils.toError(err));
        }
      },
    ]));

    // Prefix name
    $.RULE('prefixName', () => $.OR([
      {ALT: () => $.CONSUME(Sheet).image.slice(0, -1)},
      {ALT: () => $.CONSUME(SheetQuoted).image.slice(1, -2).replace(/''/g, "'")},
    ]));

    this.performSelfAnalysis();
  }
}
