import FormulaError from '../error';
import {H, Types, Criteria} from '../helpers';
import {Infix} from '../operators';


// Interface for Parser Context
interface ParserContext {
  retrieveRanges: (range: any, averageRange: any) => [any[][], any[][]];
  retrieveArg: (arg: any) => any;
  utils?: any;
}

// Interface for Statistical Functions
interface StatisticalFunctionsInterface {
  AVERAGE: (...numbers: any[]) => number;
  AVERAGEA: (...numbers: any[]) => number;
  AVERAGEIF: (context: ParserContext, range: any, criteria: any, averageRange: any) => number;
  COUNT: (...ranges: any[]) => number;
  COUNTIF: (range: any, criteria: any) => number;
}

/**
 * Statistical functions module with TypeScript type definitions
 */
const StatisticalFunctions: StatisticalFunctionsInterface = {
  /**
   * Returns the average of its arguments
   * Ignores text and logical values
   */
  AVERAGE: (...numbers): number => {
    let sum = 0, cnt = 0;

    // Use functional approach to process numbers
    H.flattenParams(numbers, Types.NUMBER, true, (item: any, info: any) => {
      if (typeof item === "number") {
        sum += item;
        cnt++;
      }
    });

    if (cnt === 0) throw FormulaError.DIV0;
    return sum / cnt;
  },

  /**
   * Returns the average of its arguments, including text and logical values
   * Text is treated as 0, TRUE as 1, FALSE as 0
   */
  AVERAGEA: (...numbers): number => {
    let sum = 0, cnt = 0;

    H.flattenParams(numbers, Types.NUMBER, true, (item: any, info: any) => {
      const type = typeof item;

      if (type === "number") {
        sum += item;
        cnt++;
      } else if (type === "boolean") {
        sum += item ? 1 : 0;
        cnt++;
      } else if (type === "string") {
        // Text is counted but treated as 0
        cnt++;
      }
    });

    if (cnt === 0) throw FormulaError.DIV0;
    return sum / cnt;
  },

  /**
   * Returns the average of all cells that meet the given criteria
   */
  AVERAGEIF: (context, range, criteria, averageRange): number => {
    const ranges = context.retrieveRanges(range, averageRange);
    range = ranges[0];
    averageRange = ranges[1];

    criteria = context.retrieveArg(criteria);
    const isCriteriaArray = criteria.isArray;
    criteria = Criteria.parse(H.accept(criteria));

    // Use reduce to calculate sum and count
    const [sum, cnt] = range.reduce((acc: [number, number], row: any[], rowNum: number) => {
      return row.reduce((innerAcc: [number, number], value: any, colNum: number) => {
        const valueToAdd = averageRange[rowNum][colNum];
        if (typeof valueToAdd !== "number") return innerAcc;

        // Check if value matches criteria
        const matches = criteria.op === 'wc'
          ? criteria.match === criteria.value.test(value)
          : Infix.compareOp(value, criteria.op, criteria.value, Array.isArray(value), isCriteriaArray);

        if (matches) {
          return [innerAcc[0] + valueToAdd, innerAcc[1] + 1];
        }
        return innerAcc;
      }, acc);
    }, [0, 0]);

    if (cnt === 0) throw FormulaError.DIV0;
    return sum / cnt;
  },

  /**
   * Counts how many numbers are in the list of arguments
   */
  COUNT: (...ranges): number => {
    // Use reduce to calculate count
    return ranges.reduce((count: number, range: any) => {
      const itemCount = typeof range === 'number' ? 1 : 0;

      if (Array.isArray(range)) {
        return count + range.flat().filter(item => typeof item === 'number').length;
      }

      if (typeof range === 'object' && range !== null) {
        H.flattenParams([range], null, true, (item: any, info: any) => {
          if ((info.isLiteral && !isNaN(item)) || typeof item === "number") {
            count++;
          }
        });
        return count;
      }

      return count + itemCount;
    }, 0);
  },

  /**
   * Counts the number of cells in a range that match the given criteria
   */
  COUNTIF: (range, criteria): number => {
    // do not flatten the array
    range = H.accept(range, Types.ARRAY, undefined, false, true);
    const isCriteriaArray = criteria.isArray;
    criteria = H.accept(criteria);

    // parse criteria
    criteria = Criteria.parse(criteria);

    // Use reduce to count matching cells
    return range.reduce((count: number, row: any[]) =>
        count + row.filter(value => {
          if (criteria.op === 'wc') {
            return criteria.match === criteria.value.test(value);
          }
          return Infix.compareOp(value, criteria.op, criteria.value, Array.isArray(value), isCriteriaArray);
        }).length
      , 0);
  }
};

export default StatisticalFunctions;
