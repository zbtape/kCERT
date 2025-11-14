export type ComplexityBand = 'Low' | 'Medium' | 'High';

/**
 * Function weights for complexity scoring.
 * Higher weights indicate more complex functions.
 * 
 * Unknown functions default to weight 2 (moderate complexity).
 * To add a new function, categorize it by complexity:
 * - Weight 1: Simple, single-purpose functions (SUM, ROUND, LEFT)
 * - Weight 2: Moderate complexity (IF, COUNTIF, PMT) - also the default for unknown functions
 * - Weight 3: High complexity (XLOOKUP, FILTER, LAMBDA, volatile functions)
 * - Weight 4: Highest complexity, volatile/indirect (OFFSET, INDIRECT)
 */
export const FUNCTION_WEIGHTS: Record<string, number> = {
    // Aggregates - simple functions
    SUM: 1,
    AVERAGE: 1,
    MIN: 1,
    MAX: 1,
    COUNT: 1,
    COUNTA: 1,
    COUNTIF: 2,
    COUNTIFS: 2,
    SUMIF: 2,
    SUMIFS: 2,
    AVERAGEIF: 2,
    AVERAGEIFS: 2,
    SUMPRODUCT: 2,
    PRODUCT: 1,
    QUOTIENT: 1,
    
    // Math & Rounding - simple functions
    ROUND: 1,
    ROUNDUP: 1,
    ROUNDDOWN: 1,
    MROUND: 1,
    CEILING: 1,
    'CEILING.MATH': 1,
    FLOOR: 1,
    'FLOOR.MATH': 1,
    TRUNC: 1,
    INT: 1,
    MOD: 1,
    ABS: 1,
    SIGN: 1,
    SQRT: 1,
    POWER: 1,
    EXP: 1,
    LN: 1,
    LOG: 1,
    LOG10: 1,
    FACT: 1,
    COMBIN: 1,
    PERMUT: 1,
    
    // Text helpers - simple string operations
    LEFT: 1,
    RIGHT: 1,
    MID: 1,
    LEN: 1,
    UPPER: 1,
    LOWER: 1,
    PROPER: 1,
    TRIM: 1,
    CLEAN: 1,
    SUBSTITUTE: 2,
    REPLACE: 2,
    FIND: 2,
    SEARCH: 2,
    TEXT: 1,
    CONCAT: 1,
    CONCATENATE: 1,
    TEXTJOIN: 2,
    VALUE: 1,
    NUMBERVALUE: 1,
    T: 1,
    EXACT: 1,
    REPT: 1,
    
    // Conditionals - moderate complexity
    IF: 2,
    IFS: 2,
    SWITCH: 2,
    AND: 1,
    OR: 1,
    NOT: 1,
    XOR: 1,
    IFERROR: 2,
    IFNA: 2,
    
    // Date/time functions
    DATE: 1,
    DATEVALUE: 1,
    EDATE: 2,
    EOMONTH: 2,
    WORKDAY: 2,
    'WORKDAY.INTL': 2,
    NETWORKDAYS: 2,
    'NETWORKDAYS.INTL': 2,
    TODAY: 3, // Volatile
    NOW: 3,   // Volatile
    YEAR: 1,
    MONTH: 1,
    DAY: 1,
    WEEKDAY: 1,
    WEEKNUM: 1,
    YEARFRAC: 2,
    DATEDIF: 2,
    TIME: 1,
    TIMEVALUE: 1,
    HOUR: 1,
    MINUTE: 1,
    SECOND: 1,
    DAYS: 1,
    DAYS360: 1,
    
    // Lookup functions - moderate to high complexity
    MATCH: 2,
    XMATCH: 2,
    INDEX: 3,
    XLOOKUP: 3,
    VLOOKUP: 3,
    HLOOKUP: 3,
    LOOKUP: 3,
    CHOOSE: 2,
    GETPIVOTDATA: 3,
    
    // Dynamic arrays - high complexity
    FILTER: 3,
    UNIQUE: 3,
    SORT: 3,
    SORTBY: 3,
    SEQUENCE: 3,
    TAKE: 3,
    DROP: 3,
    RANDARRAY: 3,
    TOCOL: 3,
    TOROW: 3,
    WRAPCOLS: 3,
    WRAPROWS: 3,
    EXPAND: 3,
    HSTACK: 3,
    VSTACK: 3,
    
    // Financial functions - moderate complexity
    PMT: 2,
    FV: 2,
    PV: 2,
    NPV: 2,
    IRR: 2,
    MIRR: 2,
    RATE: 2,
    NPER: 2,
    IPMT: 2,
    PPMT: 2,
    CUMIPMT: 2,
    CUMPRINC: 2,
    DB: 2,
    DDB: 2,
    SLN: 2,
    SYD: 2,
    VDB: 2,
    
    // Statistical functions
    STDEV: 2,
    STDEVP: 2,
    'STDEV.S': 2,
    'STDEV.P': 2,
    VAR: 2,
    VARP: 2,
    'VAR.S': 2,
    'VAR.P': 2,
    CORREL: 2,
    COVAR: 2,
    PEARSON: 2,
    RSQ: 2,
    FORECAST: 3,
    'FORECAST.LINEAR': 3,
    TREND: 3,
    GROWTH: 3,
    LINEST: 3,
    LOGEST: 3,
    PERCENTILE: 2,
    'PERCENTILE.INC': 2,
    'PERCENTILE.EXC': 2,
    QUARTILE: 2,
    'QUARTILE.INC': 2,
    'QUARTILE.EXC': 2,
    RANK: 2,
    'RANK.AVG': 2,
    'RANK.EQ': 2,
    PERCENTRANK: 2,
    'PERCENTRANK.INC': 2,
    'PERCENTRANK.EXC': 2,
    MEDIAN: 2,
    MODE: 2,
    'MODE.SNGL': 2,
    'MODE.MULT': 2,
    LARGE: 2,
    SMALL: 2,
    FREQUENCY: 2,
    
    // Structuring functions
    LET: 2,
    LAMBDA: 3,
    REDUCE: 3,
    SCAN: 3,
    MAP: 3,
    BYROW: 3,
    BYCOL: 3,
    MAKEARRAY: 3,
    
    // Volatile/indirect functions - highest complexity
    OFFSET: 4,
    INDIRECT: 4,
    RAND: 3, // Volatile
    RANDBETWEEN: 3, // Volatile
    
    // Information functions
    ISBLANK: 1,
    ISERROR: 1,
    ISNA: 1,
    ISNUMBER: 1,
    ISTEXT: 1,
    ISLOGICAL: 1,
    ISFORMULA: 1,
    ISEVEN: 1,
    ISODD: 1,
    ISREF: 1,
    N: 1,
    TYPE: 1,
    CELL: 2,
    INFO: 2,
    SHEET: 1,
    SHEETS: 1,
    
    // Array functions (legacy)
    MMULT: 3,
    TRANSPOSE: 2,
    
    // Database functions
    DSUM: 2,
    DAVERAGE: 2,
    DCOUNT: 2,
    DCOUNTA: 2,
    DMAX: 2,
    DMIN: 2,
    DPRODUCT: 2,
    DSTDEV: 2,
    DSTDEVP: 2,
    DVAR: 2,
    DVARP: 2,
    DGET: 2,
    
    // Engineering functions
    BIN2DEC: 1,
    BIN2HEX: 1,
    BIN2OCT: 1,
    DEC2BIN: 1,
    DEC2HEX: 1,
    DEC2OCT: 1,
    HEX2BIN: 1,
    HEX2DEC: 1,
    HEX2OCT: 1,
    OCT2BIN: 1,
    OCT2DEC: 1,
    OCT2HEX: 1,
    CONVERT: 2,
    
    // Reference functions
    ROW: 1,
    ROWS: 1,
    COLUMN: 1,
    COLUMNS: 1,
    AREAS: 2,
    ADDRESS: 2,
    
    // Web functions
    WEBSERVICE: 3,
    FILTERXML: 3,
    ENCODEURL: 2,
};

/**
 * F-Score thresholds for complexity band classification.
 */
export const FSCORE_THRESHOLDS = {
    medium: 30,
    high: 60,
} as const;

/**
 * Determines the complexity band based on F-Score.
 * @param score The F-Score (0-99)
 * @returns The complexity band
 */
export function getComplexityBandFromFScore(score: number): ComplexityBand {
    if (score >= FSCORE_THRESHOLDS.high) {
        return 'High';
    }
    if (score >= FSCORE_THRESHOLDS.medium) {
        return 'Medium';
    }
    return 'Low';
}

