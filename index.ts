import {FormulaParser} from './grammar/hooks';
import {DepParser} from './grammar/dependency/hooks';
import {FormulaError} from './formulas/error';
import {FormulaHelpers, Types, Factorials, WildCard, Criteria, Address} from './formulas/helpers';

// Constants
const MAX_ROW = 1048576;
const MAX_COLUMN = 16384;

// Add static properties and utility modules to FormulaParser
Object.assign(FormulaParser, {
  MAX_ROW,
  MAX_COLUMN,
  DepParser,
  FormulaError,
  FormulaHelpers,
  Types,
  Factorials,
  WildCard,
  Criteria,
  Address
});

// Export the enhanced FormulaParser as default and named export
export default FormulaParser;
export {
  FormulaParser,
  FormulaError,
  FormulaHelpers,
  Types,
  MAX_ROW,
  MAX_COLUMN,
  DepParser
};
