import {H, Types} from '../helpers';

// Interface for Engineering Functions
interface EngineeringFunctionsInterface {
  DELTA: (number1: any, number2?: any) => number;
  GESTEP: (number: any, step?: any) => number;
}

/**
 * Optimized engineering functions module with only needed functions
 * and TypeScript type definitions
 */
const EngineeringFunctions: EngineeringFunctionsInterface = {
  /**
   * Tests whether two values are equal
   * Returns 1 if number1 = number2; returns 0 otherwise
   */
  DELTA: (number1, number2 = 0): number => {
    number1 = H.accept(number1, Types.NUMBER_NO_BOOLEAN);
    number2 = H.accept(number2, Types.NUMBER_NO_BOOLEAN, 0);
    return number1 === number2 ? 1 : 0;
  },

  /**
   * Tests whether a number is greater than a threshold value
   * Returns 1 if number ≥ step; returns 0 otherwise
   */
  GESTEP: (number, step = 0): number => {
    number = H.accept(number, Types.NUMBER_NO_BOOLEAN);
    step = H.accept(step, Types.NUMBER_NO_BOOLEAN, 0);
    return number >= step ? 1 : 0;
  }
};

export default EngineeringFunctions;
