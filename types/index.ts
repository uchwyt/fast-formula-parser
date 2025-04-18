import type {TokenType} from "chevrotain";

// Parser context interface
export interface ParserContext {
  position?: {
    sheet?: string;
    row?: number;
    col?: number;
  };
  async?: boolean;
  retrieveRef: (obj: any) => any;
  getRange: (ref: any) => any;
  getCell: (ref: any) => any;

  callFunction: (name: string, args: any[]) => any;
  getVariable: (name: string) => any;
}

// Cell reference interface
export interface CellRef {
  sheet?: string;
  row: number;
  col: number;
  address?: string;
}

// Range reference interface
export interface RangeRef {
  sheet?: string;
  from: {
    row: number;
    col: number;
  };
  to: {
    row: number;
    col: number;
  };
}

// Reference object interface
export interface Reference {
  ref: CellRef | RangeRef;
}

// TokenTypes interface for vocabulary
export interface TokenTypes {
  [key: string]: TokenType;
}
