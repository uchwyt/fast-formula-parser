import FormulaError from '../error';
import {H, Types, WildCard} from '../helpers';

// Type definitions for charsets
interface CharSetDelta {
  halfRE: RegExp;
  fullRE: RegExp;
  delta: number;
}

interface CharSetMap {
  delta: number;
  half: string;
  full: string;
}

type CharSet = CharSetDelta | CharSetMap;

// Optimized charset converter
const charsets: Record<string, CharSet> = {
  latin: {halfRE: /[!-~]/g, fullRE: /[！-～]/g, delta: 0xFEE0},
  hangul1: {halfRE: /[ﾡ-ﾾ]/g, fullRE: /[ᆨ-ᇂ]/g, delta: -0xEDF9},
  hangul2: {halfRE: /[ￂ-ￜ]/g, fullRE: /[ᅡ-ᅵ]/g, delta: -0xEE61},
  kana: {
    delta: 0,
    half: "｡｢｣､･ｦｧｨｩｪｫｬｭｮｯｰｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖﾗﾘﾙﾚﾛﾜﾝﾞﾟ",
    full: "。「」、・ヲァィゥェォャュョッーアイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワン゛゜"
  },
  extras: {
    delta: 0,
    half: "¢£¬¯¦¥₩\u0020|←↑→↓■°",
    full: "￠￡￢￣￤￥￦\u3000￨￩￪￫￬￭￮"
  }
};

// Optimized character conversion functions
const toFull = (set: CharSet) => (c: string): string =>
  'delta' in set && 'halfRE' in set ?
    String.fromCharCode(c.charCodeAt(0) + set.delta) :
    [...(set as CharSetMap).full][[...(set as CharSetMap).half].indexOf(c)];

const toHalf = (set: CharSet) => (c: string): string =>
  'delta' in set && 'fullRE' in set ?
    String.fromCharCode(c.charCodeAt(0) - set.delta) :
    [...(set as CharSetMap).half][[...(set as CharSetMap).full].indexOf(c)];

const re = (set: CharSet, way: string): RegExp =>
  way === 'half' && 'halfRE' in set ? (set as CharSetDelta).halfRE :
    way === 'full' && 'fullRE' in set ? (set as CharSetDelta).fullRE :
      new RegExp(`[${(set as CharSetMap)[way as 'half' | 'full']}]`, "g");

const sets = Object.values(charsets);

// Convert to full width characters
const toFullWidth = (str: string): string =>
  sets.reduce((result, set) => result.replace(re(set, "half"), toFull(set)), str);

// Convert to half width characters
const toHalfWidth = (str: string): string =>
  sets.reduce((result, set) => result.replace(re(set, "full"), toHalf(set)), str);

/**
 * Custom number formatting functions to replace SSF dependency
 */
const NumberFormatting = {
  /**
   * Format a number with a specific number of decimal places
   * @param num The number to format
   * @param decimals Number of decimal places
   * @param withThousandsSeparator Whether to include thousands separators
   */
  formatNumber(num: number, decimals: number = 2, withThousandsSeparator: boolean = false): string {
    // Fix decimal places
    const factor = Math.pow(10, decimals);
    const fixed = Math.round(num * factor) / factor;

    // Convert to string with proper decimal places
    const parts = fixed.toFixed(decimals).split('.');
    let integerPart = parts[0];
    const decimalPart = parts.length > 1 ? parts[1] : '';

    // Add thousands separators if requested
    if (withThousandsSeparator) {
      integerPart = integerPart.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    }

    // Join parts
    return decimals === 0 ? integerPart : `${integerPart}.${decimalPart}`;
  },

  /**
   * Format a number as currency
   * @param num The number to format
   * @param decimals Number of decimal places
   * @param currencySymbol The currency symbol to use
   */
  formatCurrency(num: number, decimals: number = 2, currencySymbol: string = '$'): string {
    // Handle negative values with parentheses
    const isNegative = num < 0;
    const absValue = Math.abs(num);

    // Format with thousands separators
    const formatted = this.formatNumber(absValue, decimals, true);

    // Add symbol and handle negative values
    return isNegative ?
      `(${currencySymbol}${formatted})` :
      `${currencySymbol}${formatted}`;
  },

  /**
   * Basic implementation of Excel's TEXT function format strings
   * This handles common formats, but isn't as comprehensive as ssf
   */
  formatText(value: number, formatText: string): string {
    // Handle common Excel format codes
    if (formatText.includes('$')) {
      // Count decimal places
      const decimals = (formatText.match(/\.0+/) || [''])[0].length - 1;
      return this.formatCurrency(value, Math.max(0, decimals));
    }

    if (formatText.includes('%')) {
      // Percentage format - multiply by 100 and add % symbol
      const decimals = (formatText.match(/\.0+/) || [''])[0].length - 1;
      return this.formatNumber(value * 100, decimals) + '%';
    }

    if (formatText.includes('#,##0') || formatText.includes('0,000')) {
      // Format with thousands separators
      const decimals = (formatText.match(/\.0+/) || [''])[0].length - 1;
      return this.formatNumber(value, decimals, true);
    }

    if (formatText.match(/^0+(\.0+)?$/)) {
      // Format with leading zeros and specific decimal places
      const decimals = (formatText.match(/\.0+/) || [''])[0].length - 1;
      return this.formatNumber(value, decimals, false);
    }

    // Default - use basic number formatting
    return value.toString();
  }
};

// Interface for Text Functions
interface TextFunctionsInterface {
  CHAR: (number: any) => string;
  CLEAN: (text: any) => string;
  CODE: (text: any) => number;
  CONCAT: (...params: any[]) => string;
  CONCATENATE: (...params: any[]) => string;
  DOLLAR: (number: any, decimals?: any) => string;
  EXACT: (text1: any, text2: any) => boolean;
  FIND: (findText: any, withinText: any, startNum?: any) => number;
  LEFT: (text: any, numChars?: any) => string;
  LEN: (text: any) => number;
  LOWER: (text: any) => string;
  MID: (text: any, startNum: any, numChars: any) => string;
  NUMBERVALUE: (text: any, decimalSeparator?: any, groupSeparator?: any) => number;
  PROPER: (text: any) => string;
  REPLACE: (old_text: any, start_num: any, num_chars: any, new_text: any) => string;
  REPT: (text: any, number_times: any) => string;
  RIGHT: (text: any, numChars?: any) => string;
  SEARCH: (findText: any, withinText: any, startNum?: any) => number;
  SUBSTITUTE: (text: any, old_text: any, new_text: any, instance_num?: any) => string;
  T: (value: any) => string;
  TEXT: (value: any, formatText: any) => string;
  TRIM: (text: any) => string;
  UNICHAR: (number: any) => string;
  UNICODE: (text: any) => number;
  UPPER: (text: any) => string;
}

/**
 * Text functions module with TypeScript type definitions
 */
const TextFunctions: TextFunctionsInterface = {
  /**
   * Returns the character specified by a number
   */
  CHAR: (number): string => {
    number = H.accept(number, Types.NUMBER);
    if (number > 255 || number < 1)
      throw FormulaError.VALUE;
    return String.fromCharCode(number);
  },

  /**
   * Removes all non-printable characters from text
   */
  CLEAN: (text): string => {
    text = H.accept(text, Types.STRING);
    return text.replace(/[\x00-\x1F]/g, '');
  },

  /**
   * Returns the numeric code for the first character in a text string
   */
  CODE: (text): number => {
    text = H.accept(text, Types.STRING);
    if (text.length === 0)
      throw FormulaError.VALUE;
    return text.charCodeAt(0);
  },

  /**
   * Joins several text strings into one text string
   */
  CONCAT: (...params): string => {
    // Use modern array methods for concatenation
    const result: string[] = [];

    // does not allow union
    H.flattenParams(params, Types.STRING, false, item => {
      result.push(H.accept(item, Types.STRING));
    });

    return result.join('');
  },

  /**
   * Joins several text items into one text item
   */
  CONCATENATE: (...params): string => {
    if (params.length === 0)
      throw Error('CONCATENATE needs at least one argument.');

    // Use modern map/join for concatenation
    return params.map(param => H.accept(param, Types.STRING)).join('');
  },

  /**
   * Converts a number to text, using currency format
   */
  DOLLAR: (number, decimals = 2): string => {
    number = H.accept(number, Types.NUMBER);
    decimals = H.accept(decimals, Types.NUMBER, 2);
    return NumberFormatting.formatCurrency(number, decimals);
  },

  /**
   * Checks whether two text strings are identical
   */
  EXACT: (text1, text2): boolean => {
    text1 = H.accept(text1, Types.STRING);
    text2 = H.accept(text2, Types.STRING);
    return text1 === text2;
  },

  /**
   * Finds one text value within another (case-sensitive)
   */
  FIND: (findText, withinText, startNum = 1): number => {
    findText = H.accept(findText, Types.STRING);
    withinText = H.accept(withinText, Types.STRING);
    startNum = H.accept(startNum, Types.NUMBER, 1);

    if (startNum < 1 || startNum > withinText.length)
      throw FormulaError.VALUE;

    const res = withinText.indexOf(findText, startNum - 1);
    if (res === -1)
      throw FormulaError.VALUE;

    return res + 1;
  },

  /**
   * Returns the leftmost characters from a text value
   */
  LEFT: (text, numChars = 1): string => {
    text = H.accept(text, Types.STRING);
    numChars = H.accept(numChars, Types.NUMBER, 1);

    if (numChars < 0)
      throw FormulaError.VALUE;
    if (numChars >= text.length)
      return text;

    return text.substring(0, numChars);
  },

  /**
   * Returns the number of characters in a text string
   */
  LEN: (text): number => {
    text = H.accept(text, Types.STRING);
    return text.length;
  },

  /**
   * Converts text to lowercase
   */
  LOWER: (text): string => {
    text = H.accept(text, Types.STRING);
    return text.toLowerCase();
  },

  /**
   * Returns a specific number of characters from a text string starting at the position you specify
   */
  MID: (text, startNum, numChars): string => {
    text = H.accept(text, Types.STRING);
    startNum = H.accept(startNum, Types.NUMBER);
    numChars = H.accept(numChars, Types.NUMBER);

    if (startNum > text.length)
      return '';
    if (startNum < 1 || numChars < 0)
      throw FormulaError.VALUE;

    return text.substring(startNum - 1, startNum + numChars - 1);
  },

  /**
   * Converts text to a number in a locale-independent way
   */
  NUMBERVALUE: (text, decimalSeparator = '.', groupSeparator = ','): number => {
    text = H.accept(text, Types.STRING);
    decimalSeparator = H.accept(decimalSeparator, Types.STRING, '.');
    groupSeparator = H.accept(groupSeparator, Types.STRING, ',');

    if (text.length === 0)
      return 0;
    if (decimalSeparator.length === 0 || groupSeparator.length === 0)
      throw FormulaError.VALUE;

    // Take first character for separators
    decimalSeparator = decimalSeparator[0];
    groupSeparator = groupSeparator[0];

    if (decimalSeparator === groupSeparator ||
      text.indexOf(decimalSeparator) < text.lastIndexOf(groupSeparator))
      throw FormulaError.VALUE;

    // Use RegExp for parsing
    const res = text.replace(new RegExp(groupSeparator, 'g'), '')
      .replace(decimalSeparator, '.')
      // remove chars not related to numbers
      .replace(/[^\-0-9.%()]/g, '')
      .match(/([(-]*)([0-9]*[.]*[0-9]+)([)]?)([%]*)/);

    if (!res)
      throw FormulaError.VALUE;

    // Process number
    const [, leftParenOrMinus, numStr, rightParen, percent] = res;
    let number = Number(numStr);

    if (leftParenOrMinus.length > 1 ||
      (leftParenOrMinus.length && !rightParen.length) ||
      (!leftParenOrMinus.length && rightParen.length) ||
      isNaN(number))
      throw FormulaError.VALUE;

    // Apply percentage
    number = number / (100 ** percent.length);

    return leftParenOrMinus.length ? -number : number;
  },

  /**
   * Capitalizes the first letter in each word of a text value
   */
  PROPER: (text): string => {
    text = H.accept(text, Types.STRING);

    // Use ES-next features for better string manipulation
    return text.toLowerCase()
      .replace(/(^|\s)(\S)/g, match => match.toUpperCase());
  },

  /**
   * Replaces characters within text
   */
  REPLACE: (old_text, start_num, num_chars, new_text): string => {
    old_text = H.accept(old_text, Types.STRING);
    start_num = H.accept(start_num, Types.NUMBER);
    num_chars = H.accept(num_chars, Types.NUMBER);
    new_text = H.accept(new_text, Types.STRING);

    // Use string methods instead of array manipulation
    return old_text.substring(0, start_num - 1) +
      new_text +
      old_text.substring(start_num - 1 + num_chars);
  },

  /**
   * Repeats text a given number of times
   */
  REPT: (text, number_times): string => {
    text = H.accept(text, Types.STRING);
    number_times = H.accept(number_times, Types.NUMBER);

    // Use repeat instead of a loop
    return text.repeat(Math.max(0, Math.floor(number_times)));
  },

  /**
   * Returns the rightmost characters from a text value
   */
  RIGHT: (text, numChars = 1): string => {
    text = H.accept(text, Types.STRING);
    numChars = H.accept(numChars, Types.NUMBER, 1);

    if (numChars < 0)
      throw FormulaError.VALUE;

    const len = text.length;
    if (numChars >= len)
      return text;

    return text.substring(len - numChars);
  },

  /**
   * Finds one text value within another (not case-sensitive)
   */
  SEARCH: (findText, withinText, startNum = 1): number => {
    findText = H.accept(findText, Types.STRING);
    withinText = H.accept(withinText, Types.STRING);
    startNum = H.accept(startNum, Types.NUMBER, 1);

    if (startNum < 1 || startNum > withinText.length)
      throw FormulaError.VALUE;

    // Handle wildcards
    const findTextRegex = WildCard.isWildCard(findText)
      ? WildCard.toRegex(findText, 'i')
      : new RegExp(findText, 'i');

    const slicedText = withinText.slice(startNum - 1);
    const res = slicedText.search(findTextRegex);

    if (res === -1)
      throw FormulaError.VALUE;

    return res + startNum;
  },

  /**
   * Substitutes new text for old text in a text string
   */
  SUBSTITUTE: (text, old_text, new_text, instance_num = null): string => {
    text = H.accept(text, Types.STRING);
    old_text = H.accept(old_text, Types.STRING);
    new_text = H.accept(new_text, Types.STRING);

    // If instance_num is provided, replace only that instance
    if (instance_num !== null) {
      instance_num = H.accept(instance_num, Types.NUMBER);

      if (instance_num < 1)
        throw FormulaError.VALUE;

      const instances = text.split(old_text);

      if (instance_num > instances.length - 1)
        return text;

      let result = instances[0];
      for (let i = 1; i < instances.length; i++) {
        result += (i === instance_num ? new_text : old_text) + instances[i];
      }

      return result;
    }

    // Replace all instances
    return text.split(old_text).join(new_text);
  },

  /**
   * Returns the text referred to by value
   */
  T: (value): string => {
    value = H.accept(value);
    return typeof value === "string" ? value : '';
  },

  /**
   * Formats a number and converts it to text
   */
  TEXT: (value, formatText): string => {
    value = H.accept(value, Types.NUMBER);
    formatText = H.accept(formatText, Types.STRING);

    try {
      return NumberFormatting.formatText(value, formatText);
    } catch (e) {
      console.error(e);
      throw FormulaError.VALUE;
    }
  },

  /**
   * Removes spaces from text
   */
  TRIM: (text): string => {
    text = H.accept(text, Types.STRING);
    return text.trim();
  },

  /**
   * Returns the Unicode character that is referenced by the given numeric value
   */
  UNICHAR: (number): string => {
    number = H.accept(number, Types.NUMBER);
    if (number <= 0)
      throw FormulaError.VALUE;
    return String.fromCharCode(number);
  },

  /**
   * Returns the number (code point) corresponding to the first character of the text
   */
  UNICODE: (text): number => TextFunctions.CODE(text),

  /**
   * Converts text to uppercase
   */
  UPPER: (text): string => {
    text = H.accept(text, Types.STRING);
    return text.toUpperCase();
  },
};

export default TextFunctions;
