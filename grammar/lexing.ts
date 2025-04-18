import {createToken, Lexer, type ILexingError} from 'chevrotain';
import {FormulaError} from '../formulas/error';
import type {TokenTypes} from "../types";

// Create vocabulary dictionary
const tokenVocabulary: TokenTypes = {};

// Define tokens
const WhiteSpace = createToken({
  name: 'WhiteSpace',
  pattern: /\s+/,
  group: Lexer.SKIPPED,
});

const String = createToken({
  name: 'String',
  pattern: /"(""|[^"])*"/
});

const SingleQuotedString = createToken({
  name: 'SingleQuotedString',
  pattern: /'(''|[^'])*'/
});

const SheetQuoted = createToken({
  name: 'SheetQuoted',
  pattern: /'((?![\\\/\[\]*?:]).)+?'!/
});

const Function = createToken({
  name: 'Function',
  pattern: /[A-Za-z_]+[A-Za-z_0-9.]*\(/
});

const FormulaErrorT = createToken({
  name: 'FormulaErrorT',
  pattern: /#NULL!|#DIV\/0!|#VALUE!|#NAME\?|#NUM!|#N\/A/
});

const RefError = createToken({
  name: 'RefError',
  pattern: /#REF!/
});

const Name = createToken({
  name: 'Name',
  pattern: /[a-zA-Z_][a-zA-Z0-9_.?]*/,
});

const Sheet = createToken({
  name: 'Sheet',
  pattern: /[A-Za-z_.\d\u007F-\uFFFF]+!/
});

const Cell = createToken({
  name: 'Cell',
  pattern: /[$]?[A-Za-z]{1,3}[$]?[1-9][0-9]*/,
  longer_alt: Name
});

const Number = createToken({
  name: 'Number',
  pattern: /[0-9]+[.]?[0-9]*([eE][+\-][0-9]+)?/
});

const Boolean = createToken({
  name: 'Boolean',
  pattern: /TRUE|FALSE/i
});

const Column = createToken({
  name: 'Column',
  pattern: /[$]?[A-Za-z]{1,3}/,
  longer_alt: Name
});


/**
 * Symbols and operators
 */
const At = createToken({
  name: 'At',
  pattern: /@/
});

const Comma = createToken({
  name: 'Comma',
  pattern: /,/
});

const Colon = createToken({
  name: 'Colon',
  pattern: /:/
});

const Semicolon = createToken({
  name: 'Semicolon',
  pattern: /;/
});

const OpenParen = createToken({
  name: 'OpenParen',
  pattern: /\(/
});

const CloseParen = createToken({
  name: 'CloseParen',
  pattern: /\)/
});

const OpenSquareParen = createToken({
  name: 'OpenSquareParen',
  pattern: /\[/
});

const CloseSquareParen = createToken({
  name: 'CloseSquareParen',
  pattern: /]/
});

const ExclamationMark = createToken({
  name: 'exclamationMark',
  pattern: /!/
});

const OpenCurlyParen = createToken({
  name: 'OpenCurlyParen',
  pattern: /{/
});

const CloseCurlyParen = createToken({
  name: 'CloseCurlyParen',
  pattern: /}/
});

const QuoteS = createToken({
  name: 'QuoteS',
  pattern: /'/
});


const MulOp = createToken({
  name: 'MulOp',
  pattern: /\*/
});

const PlusOp = createToken({
  name: 'PlusOp',
  pattern: /\+/
});

const DivOp = createToken({
  name: 'DivOp',
  pattern: /\//
});

const MinOp = createToken({
  name: 'MinOp',
  pattern: /-/
});

const ConcatOp = createToken({
  name: 'ConcatOp',
  pattern: /&/
});

const ExOp = createToken({
  name: 'ExOp',
  pattern: /\^/
});

const PercentOp = createToken({
  name: 'PercentOp',
  pattern: /%/
});

const GtOp = createToken({
  name: 'GtOp',
  pattern: />/
});

const EqOp = createToken({
  name: 'EqOp',
  pattern: /=/
});

const LtOp = createToken({
  name: 'LtOp',
  pattern: /</
});

const NeqOp = createToken({
  name: 'NeqOp',
  pattern: /<>/
});

const GteOp = createToken({
  name: 'GteOp',
  pattern: />=/
});

const LteOp = createToken({
  name: 'LteOp',
  pattern: /<=/
});

// The order of tokens is important
const allTokens = [
  WhiteSpace,
  String,
  SheetQuoted,
  SingleQuotedString,
  Function,
  FormulaErrorT,
  RefError,
  Sheet,
  Cell,
  Boolean,
  Column,
  Name,
  Number,

  At,
  Comma,
  Colon,
  Semicolon,
  OpenParen,
  CloseParen,
  OpenSquareParen,
  CloseSquareParen,
  // ExclamationMark,
  OpenCurlyParen,
  CloseCurlyParen,
  QuoteS,
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
  LtOp,
];

// Export list of all token names for reference
export const allTokenNames = allTokens.map(token => token.name);

// Create lexer instance
const SelectLexer = new Lexer(allTokens, {ensureOptimizations: true});

// Populate token vocabulary
allTokens.forEach(tokenType => {
  tokenVocabulary[tokenType.name] = tokenType;
});

// Interface for lexing result with error information
export interface LexingErrorWithLocation extends ILexingError {
  message: string;
  errorLocation?: {
    line: number;
    column: number;
  }
}

// Interface for lexing result
export interface LexingResult {
  tokens: any[];
  errors: LexingErrorWithLocation[];
}

/**
 * Tokenize input text into formula tokens
 * @param inputText The formula text to tokenize
 * @returns Lexing result with tokens
 */
export const lex = (inputText: string): LexingResult => {
  const lexingResult = SelectLexer.tokenize(inputText);

  if (lexingResult.errors.length > 0) {
    const error = lexingResult.errors[0] as LexingErrorWithLocation;
    const line = error.line, column = error.column;
    let msg = '\n' + inputText.split('\n')[line - 1] + '\n';
    msg += Array(column - 1).fill(' ').join('') + '^\n';
    error.message = msg + `Error at position ${line}:${column}\n` + error.message;
    error.errorLocation = {line, column};
    throw FormulaError.ERROR(error.message, error);
  }

  return lexingResult as LexingResult;
};

export {
  tokenVocabulary,
  allTokens
};
