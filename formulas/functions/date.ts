import FormulaError from '../error';
import {H, Types} from '../helpers';

/**
 * Excel Date Functions
 * Implements date and time related functions for compatibility with Excel formulas
 */

// Constants
const MS_PER_DAY: number = 1000 * 60 * 60 * 24;
const d1900: Date = new Date(Date.UTC(1900, 0, 1));

// Weekend types mapping (used for workday calculations)
interface WeekendMap {
  [key: number]: number[];
}

const WEEKEND_TYPES: WeekendMap = {
  1: [6, 0],     // Saturday, Sunday
  2: [0, 1],     // Sunday, Monday
  3: [1, 2],
  4: [2, 3],
  5: [3, 4],
  6: [4, 5],
  7: [5, 6],
  11: [0],       // Sunday only
  12: [1],
  13: [2],
  14: [3],
  15: [4],
  16: [5],
  17: [6]        // Saturday only
};

// Week type configuration (used for WEEKDAY function)
const WEEK_TYPES: WeekendMap = {
  1: [1, 2, 3, 4, 5, 6, 7],
  2: [7, 1, 2, 3, 4, 5, 6],
  3: [6, 0, 1, 2, 3, 4, 5],
  11: [7, 1, 2, 3, 4, 5, 6],
  12: [6, 7, 1, 2, 3, 4, 5],
  13: [5, 6, 7, 1, 2, 3, 4],
  14: [4, 5, 6, 7, 1, 2, 3],
  15: [3, 4, 5, 6, 7, 1, 2],
  16: [2, 3, 4, 5, 6, 7, 1],
  17: [1, 2, 3, 4, 5, 6, 7]
};

// Week start day configuration (used for WEEKNUM)
const WEEK_STARTS: { [key: number]: number } = {
  1: 0,
  2: 1,
  12: 1,
  13: 2,
  14: 3,
  15: 4,
  16: 5,
  17: 6,
  21: 1
};

// Regular expressions for date parsing
const timeRegex = /^\s*(\d\d?)\s*(:\s*\d\d?)?\s*(:\s*\d\d?)?\s*(pm|am)?\s*$/i;
const dateRegex1 = /^\s*((\d\d?)\s*([-\/])\s*(\d\d?))([\d:.apm\s]*)$/i;
const dateRegex2 = /^\s*((\d\d?)\s*([-/])\s*(jan\w*|feb\w*|mar\w*|apr\w*|may\w*|jun\w*|jul\w*|aug\w*|sep\w*|oct\w*|nov\w*|dec\w*))([\d:.apm\s]*)$/i;
const dateRegex3 = /^\s*((jan\w*|feb\w*|mar\w*|apr\w*|may\w*|jun\w*|jul\w*|aug\w*|sep\w*|oct\w*|nov\w*|dec\w*)\s*([-/])\s*(\d\d?))([\d:.apm\s]*)$/i;

/**
 * Parse a simplified date format into a Date object
 */
function parseSimplifiedDate(text: string): Date {
  const fmt1 = text.match(dateRegex1);
  const fmt2 = text.match(dateRegex2);
  const fmt3 = text.match(dateRegex3);

  if (fmt1) {
    text = fmt1[1] + fmt1[3] + new Date().getFullYear() + fmt1[5];
  } else if (fmt2) {
    text = fmt2[1] + fmt2[3] + new Date().getFullYear() + fmt2[5];
  } else if (fmt3) {
    text = fmt3[1] + fmt3[3] + new Date().getFullYear() + fmt3[5];
  }

  return new Date(Date.parse(`${text} UTC`));
}

/**
 * Parse time string to date in UTC
 */
function parseTime(text: string): Date | undefined {
  const res = text.match(timeRegex);
  if (!res) return undefined;

  //  ["4:50:55 pm", "4", ":50", ":55", "pm", ...]
  const minutes = res[2] ? res[2] : ':00';
  const seconds = res[3] ? res[3] : ':00';
  const ampm = res[4] ? ' ' + res[4] : '';

  const date = new Date(Date.parse(`1/1/1900 ${res[1] + minutes + seconds + ampm} UTC`));
  let now = new Date();
  now = new Date(Date.UTC(now.getFullYear(), now.getMonth(), now.getDate(),
    now.getHours(), now.getMinutes(), now.getSeconds(), now.getMilliseconds()));

  return new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate(),
    date.getUTCHours(), date.getUTCMinutes(), date.getUTCSeconds(), date.getUTCMilliseconds()));
}

/**
 * Parse a UTC date to excel serial number
 * @param date - A UTC date or timestamp
 */
function toSerial(date: Date | number): number {
  if (typeof date === 'number') {
    date = new Date(date);
  }
  const timestamp = date.getTime();
  const addOn = (timestamp > -2203891200000) ? 2 : 1;
  return Math.floor((timestamp - d1900.getTime()) / MS_PER_DAY) + addOn;
}

/**
 * Parse an excel serial number to UTC date
 */
function toDate(serial: number): Date {
  if (serial < 0) {
    throw FormulaError.VALUE;
  }
  if (serial <= 60) {
    return new Date(d1900.getTime() + (serial - 1) * MS_PER_DAY);
  }
  return new Date(d1900.getTime() + (serial - 2) * MS_PER_DAY);
}

/**
 * Parse serialOrString to date and extract additional information
 */
function parseDateWithExtra(serialOrString: any): { date: Date, isDateGiven?: boolean } {
  if (serialOrString instanceof Date) return {date: serialOrString};

  serialOrString = H.accept(serialOrString);
  let isDateGiven = true;
  let date: Date;

  if (!isNaN(serialOrString)) {
    serialOrString = Number(serialOrString);
    date = toDate(serialOrString);
  } else {
    // support time without date
    date = parseTime(serialOrString);

    if (!date) {
      date = parseSimplifiedDate(serialOrString);
    } else {
      isDateGiven = false;
    }
  }

  return {date, isDateGiven};
}

/**
 * Parse any date input to a Date object
 */
function parseDate(serialOrString: any): Date {
  return parseDateWithExtra(serialOrString).date;
}

/**
 * Compare dates ignoring time component
 */
function compareDateIgnoreTime(date1: Date, date2: Date): boolean {
  return date1.getUTCFullYear() === date2.getUTCFullYear() &&
    date1.getUTCMonth() === date2.getUTCMonth() &&
    date1.getUTCDate() === date2.getUTCDate();
}

/**
 * Check if a year is a leap year
 */
function isLeapYear(year: number): boolean {
  if (year === 1900) {
    return true; // Excel bug treats 1900 as leap year
  }
  return new Date(year, 1, 29).getMonth() === 1;
}

/**
 * Date functions implementation
 */
const DateFunctions = {
  DATE: (year: any, month: any, day: any): number => {
    year = H.accept(year, Types.NUMBER);
    month = H.accept(month, Types.NUMBER);
    day = H.accept(day, Types.NUMBER);

    if (year < 0 || year >= 10000)
      throw FormulaError.NUM;

    // If year is between 0 (zero) and 1899 (inclusive), Excel adds that value to 1900
    if (year < 1900) {
      year += 1900;
    }

    return toSerial(Date.UTC(year, month - 1, day));
  },

  DATEDIF: (startDate: any, endDate: any, unit: any): number => {
    startDate = parseDate(startDate);
    endDate = parseDate(endDate);
    unit = H.accept(unit, Types.STRING).toLowerCase();

    if (startDate > endDate)
      throw FormulaError.NUM;

    const yearDiff = endDate.getUTCFullYear() - startDate.getUTCFullYear();
    const monthDiff = endDate.getUTCMonth() - startDate.getUTCMonth();
    const dayDiff = endDate.getUTCDate() - startDate.getUTCDate();
    let offset: number;

    switch (unit) {
      case 'y':
        offset = monthDiff < 0 || monthDiff === 0 && dayDiff < 0 ? -1 : 0;
        return offset + yearDiff;

      case 'm':
        offset = dayDiff < 0 ? -1 : 0;
        return yearDiff * 12 + monthDiff + offset;

      case 'd':
        return Math.floor(endDate.getTime() - startDate.getTime()) / MS_PER_DAY;

      case 'md':
        // The months and years of the dates are ignored
        startDate.setUTCFullYear(endDate.getUTCFullYear());
        if (dayDiff < 0) {
          startDate.setUTCMonth(endDate.getUTCMonth() - 1);
        } else {
          startDate.setUTCMonth(endDate.getUTCMonth());
        }
        return Math.floor(endDate.getTime() - startDate.getTime()) / MS_PER_DAY;

      case 'ym':
        // The days and years of the dates are ignored
        offset = dayDiff < 0 ? -1 : 0;
        return (offset + yearDiff * 12 + monthDiff) % 12;

      case 'yd':
        // The years of the dates are ignored
        if (monthDiff < 0 || monthDiff === 0 && dayDiff < 0) {
          startDate.setUTCFullYear(endDate.getUTCFullYear() - 1);
        } else {
          startDate.setUTCFullYear(endDate.getUTCFullYear());
        }
        return Math.floor(endDate.getTime() - startDate.getTime()) / MS_PER_DAY;

      default:
        throw FormulaError.VALUE;
    }
  },

  /**
   * Convert a date string to Excel serial number
   */
  DATEVALUE: (dateText: any): number => {
    dateText = H.accept(dateText, Types.STRING);
    const {date, isDateGiven} = parseDateWithExtra(dateText);

    if (!isDateGiven) return 0;

    const serial = toSerial(date);
    if (serial < 0 || serial > 2958465)
      throw FormulaError.VALUE;

    return serial;
  },

  DAY: (serialOrString: any): number => {
    const date = parseDate(serialOrString);
    return date.getUTCDate();
  },

  DAYS: (endDate: any, startDate: any): number => {
    endDate = parseDate(endDate);
    startDate = parseDate(startDate);

    let offset = 0;
    if (startDate.getTime() < -2203891200000 && -2203891200000 < endDate.getTime()) {
      offset = 1;
    }

    return Math.floor(endDate.getTime() - startDate.getTime()) / MS_PER_DAY + offset;
  },

  EDATE: (startDate: any, months: any): number => {
    startDate = parseDate(startDate);
    months = H.accept(months, Types.NUMBER);

    startDate.setUTCMonth(startDate.getUTCMonth() + months);
    return toSerial(startDate);
  },

  EOMONTH: (startDate: any, months: any): number => {
    startDate = parseDate(startDate);
    months = H.accept(months, Types.NUMBER);

    startDate.setUTCMonth(startDate.getUTCMonth() + months + 1, 0);
    return toSerial(startDate);
  },

  HOUR: (serialOrString: any): number => {
    const date = parseDate(serialOrString);
    return date.getUTCHours();
  },

  ISOWEEKNUM: (serialOrString: any): number => {
    const date = parseDate(serialOrString);

    // https://stackoverflow.com/questions/6117814/get-week-of-year-in-javascript-like-in-php
    const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
    const dayNum = d.getUTCDay();
    d.setUTCDate(d.getUTCDate() + 4 - dayNum);
    const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
    return Math.ceil((((d.getTime() - yearStart.getTime()) / MS_PER_DAY) + 1) / 7);
  },

  MINUTE: (serialOrString: any): number => {
    const date = parseDate(serialOrString);
    return date.getUTCMinutes();
  },

  MONTH: (serialOrString: any): number => {
    const date = parseDate(serialOrString);
    return date.getUTCMonth() + 1;
  },

  NETWORKDAYS: (startDate: any, endDate: any, holidays: any): number => {
    return DateFunctions["NETWORKDAYS.INTL"](startDate, endDate, 1, holidays);
  },

  "NETWORKDAYS.INTL": (startDate: any, endDate: any, weekend: any, holidays: any): number => {
    startDate = parseDate(startDate);
    endDate = parseDate(endDate);

    let sign = 1;
    if (startDate > endDate) {
      sign = -1;
      const temp = startDate;
      startDate = endDate;
      endDate = temp;
    }

    weekend = H.accept(weekend, null, 1);
    // Using 1111111 will always return 0.
    if (weekend === '1111111')
      return 0;

    // Parse weekend parameter
    let weekendArr: number[];

    // using weekend string, i.e, 0000011
    if (typeof weekend === "string" && Number(weekend).toString() !== weekend) {
      if (weekend.length !== 7) throw FormulaError.VALUE;

      weekend = weekend.charAt(6) + weekend.slice(0, 6);
      weekendArr = [];

      for (let i = 0; i < weekend.length; i++) {
        if (weekend.charAt(i) === '1')
          weekendArr.push(i);
      }
    } else {
      // using weekend number
      if (typeof weekend !== "number")
        throw FormulaError.VALUE;

      weekendArr = WEEKEND_TYPES[weekend];

      if (!weekendArr)
        throw FormulaError.NUM;
    }

    // Parse holidays
    const holidaysArr: Date[] = [];
    if (holidays != null) {
      H.flattenParams([holidays], Types.NUMBER, false, (item: any) => {
        holidaysArr.push(parseDate(item));
      });
    }

    let numWorkDays = 0;
    const current = new Date(startDate.getTime());

    while (current <= endDate) {
      // Check if current day is a weekend
      let isWeekend = false;
      for (let i = 0; i < weekendArr.length; i++) {
        if (weekendArr[i] === current.getUTCDay()) {
          isWeekend = true;
          break;
        }
      }

      if (!isWeekend) {
        // Check if current day is a holiday
        let isHoliday = false;
        for (let i = 0; i < holidaysArr.length; i++) {
          if (compareDateIgnoreTime(current, holidaysArr[i])) {
            isHoliday = true;
            break;
          }
        }

        // Count as work day if not a holiday
        if (!isHoliday) {
          numWorkDays++;
        }
      }

      // Move to next day
      current.setUTCDate(current.getUTCDate() + 1);
    }

    return sign * numWorkDays;
  },

  NOW: (): number => {
    const now = new Date();
    return toSerial(Date.UTC(now.getFullYear(), now.getMonth(), now.getDate(),
        now.getHours(), now.getMinutes(), now.getSeconds(), now.getMilliseconds()))
      + (3600 * now.getHours() + 60 * now.getMinutes() + now.getSeconds()) / 86400;
  },

  SECOND: (serialOrString: any): number => {
    const date = parseDate(serialOrString);
    return date.getUTCSeconds();
  },

  TIME: (hour: any, minute: any, second: any): number => {
    hour = H.accept(hour, Types.NUMBER);
    minute = H.accept(minute, Types.NUMBER);
    second = H.accept(second, Types.NUMBER);

    if (hour < 0 || hour > 32767 || minute < 0 || minute > 32767 || second < 0 || second > 32767)
      throw FormulaError.NUM;

    return (3600 * hour + 60 * minute + second) / 86400;
  },

  TIMEVALUE: (timeText: any): number => {
    const date = parseDate(timeText);
    return (3600 * date.getUTCHours() + 60 * date.getUTCMinutes() + date.getUTCSeconds()) / 86400;
  },

  TODAY: (): number => {
    const now = new Date();
    return toSerial(Date.UTC(now.getFullYear(), now.getMonth(), now.getDate()));
  },

  WEEKDAY: (serialOrString: any, returnType: any = 1): number => {
    const date = parseDate(serialOrString);
    returnType = H.accept(returnType, Types.NUMBER, 1);

    const day = date.getUTCDay();
    const weekTypes = WEEK_TYPES[returnType];

    if (!weekTypes)
      throw FormulaError.NUM;

    return weekTypes[day];
  },

  WEEKNUM: (serialOrString: any, returnType: any = 1): number => {
    const date = parseDate(serialOrString);
    returnType = H.accept(returnType, Types.NUMBER, 1);

    if (returnType === 21) {
      return DateFunctions.ISOWEEKNUM(serialOrString);
    }

    const weekStart = WEEK_STARTS[returnType];

    if (weekStart === undefined)
      throw FormulaError.NUM;

    const yearStart = new Date(Date.UTC(date.getUTCFullYear(), 0, 1));
    const offset = yearStart.getUTCDay() < weekStart ? 1 : 0;

    return Math.ceil((((date.getTime() - yearStart.getTime()) / MS_PER_DAY) + 1) / 7) + offset;
  },

  WORKDAY: (startDate: any, days: any, holidays: any): number => {
    return DateFunctions["WORKDAY.INTL"](startDate, days, 1, holidays);
  },

  "WORKDAY.INTL": (startDate: any, days: any, weekend: any, holidays: any): number => {
    startDate = parseDate(startDate);
    days = H.accept(days, Types.NUMBER);

    weekend = H.accept(weekend, null, 1);
    // Using 1111111 will always return value error.
    if (weekend === '1111111')
      throw FormulaError.VALUE;

    // Parse weekend parameter
    let weekendArr: number[];

    // using weekend string, i.e, 0000011
    if (typeof weekend === "string" && Number(weekend).toString() !== weekend) {
      if (weekend.length !== 7)
        throw FormulaError.VALUE;

      weekend = weekend.charAt(6) + weekend.slice(0, 6);
      weekendArr = [];

      for (let i = 0; i < weekend.length; i++) {
        if (weekend.charAt(i) === '1')
          weekendArr.push(i);
      }
    } else {
      // using weekend number
      if (typeof weekend !== "number")
        throw FormulaError.VALUE;

      weekendArr = WEEKEND_TYPES[weekend];

      if (!weekendArr)
        throw FormulaError.NUM;
    }

    // Parse holidays
    const holidaysArr: Date[] = [];
    if (holidays != null) {
      H.flattenParams([holidays], Types.NUMBER, false, (item: any) => {
        holidaysArr.push(parseDate(item));
      });
    }

    // Move one day forward to start counting
    startDate.setUTCDate(startDate.getUTCDate() + 1);
    let cnt = 0;

    while (cnt < days) {
      let skip = false;

      // Check if current day is a weekend
      for (let i = 0; i < weekendArr.length; i++) {
        if (weekendArr[i] === startDate.getUTCDay()) {
          skip = true;
          break;
        }
      }

      if (!skip) {
        // Check if current day is a holiday
        let found = false;
        for (let i = 0; i < holidaysArr.length; i++) {
          if (compareDateIgnoreTime(startDate, holidaysArr[i])) {
            found = true;
            break;
          }
        }

        if (!found) cnt++;
      }

      startDate.setUTCDate(startDate.getUTCDate() + 1);
    }

    return toSerial(startDate) - 1;
  },

  YEAR: (serialOrString: any): number => {
    const date = parseDate(serialOrString);
    return date.getUTCFullYear();
  },

  // Warning: complex implementation with multiple calculation methods
  YEARFRAC: (startDate: any, endDate: any, basis: any = 0): number => {
    startDate = parseDate(startDate);
    endDate = parseDate(endDate);

    if (startDate > endDate) {
      const temp = startDate;
      startDate = endDate;
      endDate = temp;
    }

    basis = H.accept(basis, Types.NUMBER, 0);
    basis = Math.trunc(basis);

    if (basis < 0 || basis > 4)
      throw FormulaError.VALUE;

    // Extract date components for calculations
    const sd = startDate.getUTCDate();
    const sm = startDate.getUTCMonth() + 1;
    const sy = startDate.getUTCFullYear();
    const ed = endDate.getUTCDate();
    const em = endDate.getUTCMonth() + 1;
    const ey = endDate.getUTCFullYear();

    switch (basis) {
      case 0:
        // US (NASD) 30/360
        if (sd === 31 && ed === 31) {
          startDate.setUTCDate(30);
          endDate.setUTCDate(30);
        } else if (sd === 31) {
          startDate.setUTCDate(30);
        } else if (sd === 30 && ed === 31) {
          endDate.setUTCDate(30);
        }

        return Math.abs(
          (ed + em * 30 + ey * 360) - (sd + sm * 30 + sy * 360)
        ) / 360;

      case 1:
        // Actual/actual
        if (ey - sy < 2) {
          const yLength = isLeapYear(sy) && sy !== 1900 ? 366 : 365;
          const days = DateFunctions.DAYS(endDate, startDate);
          return days / yLength;
        } else {
          const years = (ey - sy) + 1;
          const days = (new Date(ey + 1, 0, 1).getTime() - new Date(sy, 0, 1).getTime()) / 1000 / 60 / 60 / 24;
          const average = days / years;
          return DateFunctions.DAYS(endDate, startDate) / average;
        }

      case 2:
        // Actual/360
        return Math.abs(DateFunctions.DAYS(endDate, startDate) / 360);

      case 3:
        // Actual/365
        return Math.abs(DateFunctions.DAYS(endDate, startDate) / 365);

      case 4:
        // European 30/360
        return Math.abs(
          (ed + em * 30 + ey * 360) - (sd + sm * 30 + sy * 360)
        ) / 360;
    }

    return 0; // Should never reach here
  },
};

export default DateFunctions;
