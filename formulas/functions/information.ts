import {FormulaError} from '../error';
import {H, Types} from '../helpers';

// Interface for Information Functions
interface InfoFunctionsInterface {
  ISBLANK: (value: any) => boolean;
  ISERR: (value: any) => boolean;
  ISERROR: (value: any) => boolean;
  ISEVEN: (number: any) => boolean;
  ISLOGICAL: (value: any) => boolean;
  ISNA: (value: any) => boolean;
  ISNONTEXT: (value: any) => boolean;
  ISNUMBER: (value: any) => boolean;
  ISTEXT: (value: any) => boolean;
}

/**
 * Optimized information functions module with TypeScript type definitions
 */
const InfoFunctions: InfoFunctionsInterface = {
  /**
   * Checks if a value is blank
   * @param value - Value to check
   * @returns true if value is blank
   */
  ISBLANK: (value): boolean =>
    value?.ref ? value.value == null || value.value === '' : false,

  /**
   * Checks if a value is an error other than #N/A
   * @param value - Value to check
   * @returns true if value is an error other than #N/A
   */
  ISERR: (value): boolean => {
    value = H.accept(value);
    return value instanceof FormulaError && value !== FormulaError.NA;
  },

  /**
   * Checks if a value is any error
   * @param value - Value to check
   * @returns true if value is any error
   */
  ISERROR: (value): boolean =>
    H.accept(value) instanceof FormulaError,

  /**
   * Checks if a number is even
   * @param number - Number to check
   * @returns true if number is even
   */
  ISEVEN: (number): boolean => {
    number = H.accept(number, Types.NUMBER);
    return (Math.trunc(number) & 1) === 0; // Bitwise check for even
  },

  /**
   * Checks if a value is a logical value (TRUE or FALSE)
   * @param value - Value to check
   * @returns true if value is a logical value
   */
  ISLOGICAL: (value): boolean =>
    typeof H.accept(value) === 'boolean',

  /**
   * Checks if a value is the #N/A error
   * @param value - Value to check
   * @returns true if value is #N/A
   */
  ISNA: (value): boolean =>
    H.accept(value) === FormulaError.NA,

  /**
   * Checks if a value is not text
   * @param value - Value to check
   * @returns true if value is not text
   */
  ISNONTEXT: (value): boolean =>
    typeof H.accept(value) !== 'string',

  /**
   * Checks if a value is a number
   * @param value - Value to check
   * @returns true if value is a number
   */
  ISNUMBER: (value): boolean =>
    typeof H.accept(value) === "number",

  /**
   * Checks if a value is text
   * @param value - Value to check
   * @returns true if value is text
   */
  ISTEXT: (value): boolean =>
    typeof H.accept(value) === 'string',
};

export default InfoFunctions;
