#!/usr/bin/perl -w
use warnings;
use strict;

package Spreadsheet::Nifty::XLS::Formula;

use constant
{
  TOKEN_CLASS_BASIC     => 0x0,
  TOKEN_CLASS_REFERENCE => 0x1,
  TOKEN_CLASS_VALUE     => 0x2,
  TOKEN_CLASS_ARRAY     => 0x3,
};

use constant
{
  TOKEN_MASK_ID    => 0x1F,

  TOKEN_MASK_CLASS => 0x60,
  TOKEN_SHIFT_CLASS => 5,
};

my $functions =
[
  {id =>  0, name => 'COUNT',     minArgs => 0, maxArgs => 30},
  {id =>  1, name => 'IF',        minArgs => 2, maxArgs => 3},
  {id =>  2, name => 'ISNA',      minArgs => 1, maxArgs => 1},
  {id =>  3, name => 'ISERROR',   minArgs => 1, maxArgs => 1},
  {id =>  4, name => 'SUM',       minArgs => 0, maxArgs => 30},
  {id =>  5, name => 'AVERAGE',   minArgs => 1, maxArgs => 30},
  {id =>  6, name => 'MIN',       minArgs => 1, maxArgs => 30},
  {id =>  7, name => 'MAX',       minArgs => 1, maxArgs => 30},
  {id =>  8, name => 'ROW',       minArgs => 0, maxArgs => 1},
  {id =>  9, name => 'COLUMN',    minArgs => 0, maxArgs => 1},
  {id => 10, name => 'NA',        minArgs => 0, maxArgs => 0},
  {id => 11, name => 'NPV',       minArgs => 2, maxArgs => 30},
  {id => 12, name => 'STDEV',     minArgs => 1, maxArgs => 30},
  {id => 13, name => 'DOLLAR',    minArgs => 1, maxArgs => 2},
  {id => 14, name => 'FIXED',     minArgs => 2, maxArgs => 3},
  {id => 15, name => 'SIN',       minArgs => 1, maxArgs => 1},
  {id => 16, name => 'COS',       minArgs => 1, maxArgs => 1},
  {id => 17, name => 'TAN',       minArgs => 1, maxArgs => 1},
  {id => 18, name => 'ATAN',      minArgs => 1, maxArgs => 1},
  {id => 19, name => 'PI',        minArgs => 0, maxArgs => 0},
  {id => 20, name => 'SQRT',      minArgs => 1, maxArgs => 1},
  {id => 21, name => 'EXP',       minArgs => 1, maxArgs => 1},
  {id => 22, name => 'LN',        minArgs => 1, maxArgs => 1},
  {id => 23, name => 'LOG10',     minArgs => 1, maxArgs => 1},
  {id => 24, name => 'ABS',       minArgs => 1, maxArgs => 1},
  {id => 25, name => 'INT',       minArgs => 1, maxArgs => 1},
  {id => 26, name => 'SIGN',      minArgs => 1, maxArgs => 1},
  {id => 27, name => 'ROUND',     minArgs => 2, maxArgs => 2},
  {id => 28, name => 'LOOKUP',    minArgs => 2, maxArgs => 3},
  {id => 29, name => 'INDEX',     minArgs => 2, maxArgs => 4},
  {id => 30, name => 'REPT',      minArgs => 2, maxArgs => 2},
  {id => 31, name => 'MID',       minArgs => 3, maxArgs => 3},
  {id => 32, name => 'LEN',       minArgs => 1, maxArgs => 1},
  {id => 33, name => 'VALUE',     minArgs => 1, maxArgs => 1},
  {id => 34, name => 'TRUE',      minArgs => 0, maxArgs => 0},
  {id => 35, name => 'FALSE',     minArgs => 0, maxArgs => 0},
  {id => 36, name => 'AND',       minArgs => 1, maxArgs => 30},
  {id => 37, name => 'OR',        minArgs => 1, maxArgs => 30},
  {id => 38, name => 'NOT',       minArgs => 1, maxArgs => 1},
  {id => 39, name => 'MOD',       minArgs => 2, maxArgs => 2},
  {id => 40, name => 'DCOUNT',    minArgs => 3, maxArgs => 3},
  {id => 41, name => 'DSUM',      minArgs => 3, maxArgs => 3},
  {id => 42, name => 'DAVERAGE',  minArgs => 3, maxArgs => 3},
  {id => 43, name => 'DMIN',      minArgs => 3, maxArgs => 3},
  {id => 44, name => 'DMAX',      minArgs => 3, maxArgs => 3},
  {id => 45, name => 'DSTDEV',    minArgs => 3, maxArgs => 3},
  {id => 46, name => 'VAR',       minArgs => 1, maxArgs => 30},
  {id => 47, name => 'DVAR',      minArgs => 3, maxArgs => 3},
  {id => 48, name => 'TEXT',      minArgs => 2, maxArgs => 2},
  {id => 49, name => 'LINEST',    minArgs => 1, maxArgs => 4},
  {id => 50, name => 'TREND',     minArgs => 1, maxArgs => 4},
  {id => 51, name => 'LOGEST',    minArgs => 1, maxArgs => 4},
  {id => 52, name => 'GROWTH',    minArgs => 1, maxArgs => 4},
  {id => 56, name => 'PV',        minArgs => 3, maxArgs => 5},
  {id => 57, name => 'FV',        minArgs => 3, maxArgs => 5},
  {id => 58, name => 'NPER',      minArgs => 3, maxArgs => 5},
  {id => 59, name => 'PMT',       minArgs => 3, maxArgs => 5},
  {id => 60, name => 'RATE',      minArgs => 3, maxArgs => 6},
  {id => 61, name => 'MIRR',      minArgs => 3, maxArgs => 3},
  {id => 62, name => 'IRR',       minArgs => 1, maxArgs => 2},
  {id => 63, name => 'RAND',      minArgs => 0, maxArgs => 0},
  {id => 64, name => 'MATCH',     minArgs => 2, maxArgs => 3},
  {id => 65, name => 'DATE',      minArgs => 3, maxArgs => 3},
  {id => 66, name => 'TIME',      minArgs => 3, maxArgs => 3},

  {id => 67, name => 'DAY',       minArgs => 1, maxArgs => 1},
  {id => 68, name => 'MONTH',     minArgs => 1, maxArgs => 1},
  {id => 69, name => 'YEAR',      minArgs => 1, maxArgs => 1},
  {id => 70, name => 'WEEKDAY',   minArgs => 1, maxArgs => 2},
  {id => 71, name => 'HOUR',      minArgs => 1, maxArgs => 1},
  {id => 72, name => 'MINUTE',    minArgs => 1, maxArgs => 1},
  {id => 73, name => 'SECOND',    minArgs => 1, maxArgs => 1},
  {id => 74, name => 'NOW',       minArgs => 0, maxArgs => 0},
  {id => 75, name => 'AREAS',     minArgs => 1, maxArgs => 1},
  {id => 76, name => 'ROWS',      minArgs => 1, maxArgs => 1},
  {id => 77, name => 'COLUMNS',   minArgs => 1, maxArgs => 1},
  {id => 78, name => 'OFFSET',    minArgs => 3, maxArgs => 5},
  {id => 82, name => 'SEARCH',    minArgs => 2, maxArgs => 3},
  {id => 83, name => 'TRANSPOSE', minArgs => 1, maxArgs => 1},
  {id => 86, name => 'TYPE',      minArgs => 1, maxArgs => 1},
  {id => 97, name => 'ATAN2',     minArgs => 2, maxArgs => 2},
  {id => 98, name => 'ASIN',      minArgs => 1, maxArgs => 1},
  {id => 99, name => 'ACOS',      minArgs => 1, maxArgs => 1},
  {id => 100, name => 'CHOOSE',   minArgs => 2, maxArgs => 30},
  {id => 101, name => 'HLOOKUP',  minArgs => 3, maxArgs => 4},
  {id => 102, name => 'VLOOKUP',  minArgs => 3, maxArgs => 4},
  {id => 105, name => 'ISREF',    minArgs => 1, maxArgs => 1},
  {id => 109, name => 'LOG',      minArgs => 1, maxArgs => 2},
  {id => 111, name => 'CHAR',     minArgs => 1, maxArgs => 1},
  {id => 112, name => 'LOWER',    minArgs => 1, maxArgs => 1},
  {id => 113, name => 'UPPER',    minArgs => 1, maxArgs => 1},
  {id => 114, name => 'PROPER',   minArgs => 1, maxArgs => 1},
  {id => 115, name => 'LEFT',     minArgs => 1, maxArgs => 2},


  {id => 116, name => 'RIGHT',      minArgs => 1, maxArgs => 2},
  {id => 117, name => 'EXACT',      minArgs => 2, maxArgs => 2},
  {id => 118, name => 'TRIM',       minArgs => 1, maxArgs => 1},
  {id => 119, name => 'REPLACE',    minArgs => 4, maxArgs => 4},
  {id => 120, name => 'SUBSTITUTE', minArgs => 3, maxArgs => 4},
  {id => 121, name => 'CODE',       minArgs => 1, maxArgs => 1},
  {id => 124, name => 'FIND',       minArgs => 2, maxArgs => 3},
  {id => 125, name => 'CELL',       minArgs => 1, maxArgs => 2},
  {id => 126, name => 'ISERR',      minArgs => 1, maxArgs => 1},
  {id => 127, name => 'ISTEXT',     minArgs => 1, maxArgs => 1},
  {id => 128, name => 'ISNUMBER',   minArgs => 1, maxArgs => 1},
  {id => 129, name => 'ISBLANK',    minArgs => 1, maxArgs => 1},
  {id => 130, name => 'T',          minArgs => 1, maxArgs => 1},
  {id => 131, name => 'N',          minArgs => 1, maxArgs => 1},
  {id => 140, name => 'DATEVALUE',  minArgs => 1, maxArgs => 1},
  {id => 141, name => 'TIMEVALUE',  minArgs => 1, maxArgs => 1},
  {id => 142, name => 'SLN',        minArgs => 3, maxArgs => 3},
  {id => 143, name => 'SYD',        minArgs => 4, maxArgs => 4},
  {id => 144, name => 'DDB',        minArgs => 4, maxArgs => 5},
  {id => 148, name => 'INDIRECT',   minArgs => 1, maxArgs => 2},
  {id => 162, name => 'CLEAN',      minArgs => 1, maxArgs => 1},
  {id => 163, name => 'MDETERM',    minArgs => 1, maxArgs => 1},
  {id => 164, name => 'MINVERSE',   minArgs => 1, maxArgs => 1},
  {id => 165, name => 'MMULT',      minArgs => 2, maxArgs => 2},
  {id => 167, name => 'IPMT',       minArgs => 4, maxArgs => 6},
  {id => 168, name => 'PPMT',       minArgs => 4, maxArgs => 6},
  {id => 169, name => 'COUNTA',     minArgs => 0, maxArgs => 30},
  {id => 183, name => 'PRODUCT',    minArgs => 0, maxArgs => 30},
  {id => 184, name => 'FACT',       minArgs => 1, maxArgs => 1},
  {id => 189, name => 'DPRODUCT',   minArgs => 3, maxArgs => 3},
  {id => 190, name => 'ISNONTEXT',  minArgs => 1, maxArgs => 1},
  {id => 193, name => 'STDEVP',     minArgs => 1, maxArgs => 30},
  {id => 194, name => 'VARP',       minArgs => 1, maxArgs => 30},
  {id => 195, name => 'DSTDEVP',    minArgs => 3, maxArgs => 3},
  {id => 196, name => 'DVARP',      minArgs => 3, maxArgs => 3},
  {id => 197, name => 'TRUNC',      minArgs => 1, maxArgs => 2},
  {id => 198, name => 'ISLOGICAL',  minArgs => 1, maxArgs => 1},
  {id => 199, name => 'DCOUNTA',    minArgs => 3, maxArgs => 3},

  {id => 204, name => 'USDOLLAR',   minArgs => 1, maxArgs => 2},
  {id => 205, name => 'FINDB',      minArgs => 2, maxArgs => 3},
  {id => 206, name => 'SEARCHB',    minArgs => 2, maxArgs => 3},
  {id => 207, name => 'REPLACEB',   minArgs => 4, maxArgs => 4},
  {id => 208, name => 'LEFTB',      minArgs => 1, maxArgs => 2},
  {id => 209, name => 'RIGHTB',     minArgs => 1, maxArgs => 2},
  {id => 210, name => 'MIDB',       minArgs => 3, maxArgs => 3},
  {id => 211, name => 'LENB',       minArgs => 1, maxArgs => 1},
  {id => 212, name => 'ROUNDUP',    minArgs => 2, maxArgs => 2},
  {id => 213, name => 'ROUNDDOWN',  minArgs => 2, maxArgs => 2},
  {id => 214, name => 'ASC',        minArgs => 1, maxArgs => 1},

  {id => 215, name => 'DBCS',       minArgs => 1, maxArgs => 1},
  {id => 219, name => 'ADDRESS',    minArgs => 2, maxArgs => 5},
  {id => 220, name => 'DAYS360',    minArgs => 2, maxArgs => 3},
  {id => 221, name => 'TODAY',      minArgs => 0, maxArgs => 0},
  {id => 222, name => 'VDB',        minArgs => 5, maxArgs => 7},
  {id => 227, name => 'MEDIAN',     minArgs => 1, maxArgs => 30},
  {id => 228, name => 'SUMPRODUCT', minArgs => 1, maxArgs => 30},
  {id => 229, name => 'SINH',       minArgs => 1, maxArgs => 1},
  {id => 230, name => 'COSH',       minArgs => 1, maxArgs => 1},
  {id => 231, name => 'TANH',       minArgs => 1, maxArgs => 1},
  {id => 232, name => 'ASINH',      minArgs => 1, maxArgs => 1},
  {id => 233, name => 'ACOSH',      minArgs => 1, maxArgs => 1},
  {id => 234, name => 'ATANH',      minArgs => 1, maxArgs => 1},
  {id => 235, name => 'DGET',       minArgs => 3, maxArgs => 3},
  {id => 244, name => 'INFO',       minArgs => 1, maxArgs => 1},

  {id => 216, name => 'RANK',       minArgs => 2, maxArgs => 3},
  {id => 247, name => 'DB',         minArgs => 4, maxArgs => 5},
  {id => 252, name => 'FREQUENCY',  minArgs => 2, maxArgs => 2},

  # Function id 255 is used for user-defined or future functions

  {id => 261, name => 'ERROR.TYPE',   minArgs => 1, maxArgs => 1},
  {id => 269, name => 'AVEDEV',       minArgs => 1, maxArgs => 30},
  {id => 270, name => 'BETADIST',     minArgs => 3, maxArgs => 5},
  {id => 271, name => 'GAMMALN',      minArgs => 1, maxArgs => 1},
  {id => 272, name => 'BETAINV',      minArgs => 3, maxArgs => 5},
  {id => 273, name => 'BINOMDIST',    minArgs => 4, maxArgs => 4},
  {id => 274, name => 'CHIDIST',      minArgs => 2, maxArgs => 2},
  {id => 275, name => 'CHIINV',       minArgs => 2, maxArgs => 2},
  {id => 276, name => 'COMBIN',       minArgs => 2, maxArgs => 2},
  {id => 277, name => 'CONFIDENCE',   minArgs => 3, maxArgs => 3},
  {id => 278, name => 'CRITBINOM',    minArgs => 3, maxArgs => 3},
  {id => 279, name => 'EVEN',         minArgs => 1, maxArgs => 1},
  {id => 280, name => 'EXPONDIST',    minArgs => 3, maxArgs => 3},
  {id => 281, name => 'FDIST',        minArgs => 3, maxArgs => 3},
  {id => 282, name => 'FINV',         minArgs => 3, maxArgs => 3},
  {id => 283, name => 'FISHER',       minArgs => 1, maxArgs => 1},
  {id => 284, name => 'FISHERINV',    minArgs => 1, maxArgs => 1},
  {id => 285, name => 'FLOOR',        minArgs => 2, maxArgs => 2},
  {id => 286, name => 'GAMMADIST',    minArgs => 4, maxArgs => 4},
  {id => 287, name => 'GAMMAINV',     minArgs => 3, maxArgs => 3},
  {id => 288, name => 'CEILING',      minArgs => 2, maxArgs => 2},
  {id => 289, name => 'HYPGEOMDIST',  minArgs => 4, maxArgs => 4},
  {id => 290, name => 'LOGNORMDIST',  minArgs => 3, maxArgs => 3},
  {id => 291, name => 'LOGINV',       minArgs => 3, maxArgs => 3},
  {id => 292, name => 'NEGBINOMDIST', minArgs => 3, maxArgs => 3},
  {id => 293, name => 'NORMDIST',     minArgs => 4, maxArgs => 4},
  {id => 294, name => 'NORMSDIST',    minArgs => 1, maxArgs => 1},
  {id => 295, name => 'NORMINV',      minArgs => 3, maxArgs => 3},
  {id => 296, name => 'NORMSINV',     minArgs => 1, maxArgs => 1},
  {id => 297, name => 'STANDARDIZE',  minArgs => 3, maxArgs => 3},
  {id => 298, name => 'ODD',          minArgs => 1, maxArgs => 1},
  {id => 299, name => 'PERMUT',       minArgs => 2, maxArgs => 2},
  {id => 300, name => 'POISSON',      minArgs => 3, maxArgs => 3},
  {id => 301, name => 'TDIST',        minArgs => 3, maxArgs => 3},
  {id => 302, name => 'WEIBULL',      minArgs => 4, maxArgs => 4},
  {id => 303, name => 'SUMXMY2',      minArgs => 2, maxArgs => 2},
  {id => 304, name => 'SUMX2MY2',     minArgs => 2, maxArgs => 2},
  {id => 305, name => 'SUMX2PY2',     minArgs => 2, maxArgs => 2},
  {id => 306, name => 'CHITEST',      minArgs => 2, maxArgs => 2},
  {id => 307, name => 'CORREL',       minArgs => 2, maxArgs => 2},
  {id => 308, name => 'COVAR',        minArgs => 2, maxArgs => 2},
  {id => 309, name => 'FORECAST',     minArgs => 3, maxArgs => 3},
  {id => 310, name => 'FTEST',        minArgs => 2, maxArgs => 2},
  {id => 311, name => 'INTERCEPT',    minArgs => 2, maxArgs => 2},
  {id => 312, name => 'PEARSON',      minArgs => 2, maxArgs => 2},
  {id => 313, name => 'RSQ',          minArgs => 2, maxArgs => 2},
  {id => 314, name => 'STEYX',        minArgs => 2, maxArgs => 2},
  {id => 315, name => 'SLOPE',        minArgs => 2, maxArgs => 2},
  {id => 316, name => 'TTEST',        minArgs => 4, maxArgs => 4},
  {id => 317, name => 'PROB',         minArgs => 3, maxArgs => 4},
  {id => 318, name => 'DEVSQ',        minArgs => 1, maxArgs => 30},
  {id => 319, name => 'GEOMEAN',      minArgs => 1, maxArgs => 30},
  {id => 320, name => 'HARMEAN',      minArgs => 1, maxArgs => 30},
  {id => 321, name => 'SUMSQ',        minArgs => 0, maxArgs => 30},
  {id => 322, name => 'KURT',         minArgs => 1, maxArgs => 30},
  {id => 323, name => 'SKEW',         minArgs => 1, maxArgs => 30},
  {id => 324, name => 'ZTEST',        minArgs => 2, maxArgs => 3},
  {id => 325, name => 'LARGE',        minArgs => 2, maxArgs => 2},
  {id => 326, name => 'SMALL',        minArgs => 2, maxArgs => 2},
  {id => 327, name => 'QUARTILE',     minArgs => 2, maxArgs => 2},
  {id => 328, name => 'PERCENTILE',   minArgs => 2, maxArgs => 2},
  {id => 329, name => 'PERCENTRANK',  minArgs => 2, maxArgs => 3},
  {id => 330, name => 'MODE',         minArgs => 1, maxArgs => 30},
  {id => 331, name => 'TRIMMEAN',     minArgs => 2, maxArgs => 2},
  {id => 332, name => 'TINV',         minArgs => 2, maxArgs => 2},

  {id => 336, name => 'CONCATENATE', minArgs => 0, maxArgs => 30},
  {id => 337, name => 'POWER',       minArgs => 2, maxArgs => 2},
  {id => 342, name => 'RADIANS',     minArgs => 1, maxArgs => 1},
  {id => 343, name => 'DEGREES',     minArgs => 1, maxArgs => 1},
  {id => 344, name => 'SUBTOTAL',    minArgs => 2, maxArgs => 30},

  {id => 345, name => 'SUMIF',        minArgs => 2, maxArgs => 3},
  {id => 346, name => 'COUNTIF',      minArgs => 2, maxArgs => 2},
  {id => 347, name => 'COUNTBLANK',   minArgs => 1, maxArgs => 1},
  {id => 350, name => 'ISPMT',        minArgs => 4, maxArgs => 4},
  {id => 351, name => 'DATEDIF',      minArgs => 3, maxArgs => 3},
  {id => 352, name => 'DATESTRING',   minArgs => 1, maxArgs => 1},
  {id => 353, name => 'NUMBERSTRING', minArgs => 2, maxArgs => 2},
  {id => 354, name => 'ROMAN',        minArgs => 1, maxArgs => 2},

  {id => 358, name => 'GETPIVOTDATA', minArgs => 2, maxArgs => 30},
  {id => 359, name => 'HYPERLINK',    minArgs => 1, maxArgs => 2},
  {id => 360, name => 'PHONETIC',     minArgs => 1, maxArgs => 1},
  {id => 361, name => 'AVERAGEA',     minArgs => 1, maxArgs => 30},
  {id => 362, name => 'MAXA',         minArgs => 1, maxArgs => 30},

  {id => 363, name => 'MINA',    minArgs => 1, maxArgs => 30},
  {id => 364, name => 'STDEVPA', minArgs => 1, maxArgs => 30},
  {id => 365, name => 'VARPA',   minArgs => 1, maxArgs => 30},
  {id => 366, name => 'STDEVA',  minArgs => 1, maxArgs => 30},
  {id => 367, name => 'VARA',    minArgs => 1, maxArgs => 30},
];

my $basicTokens =
[
  {id => 0x01, name => 'tExp', size5 => 5, size8 => 5},
  {id => 0x02, name => 'tTbl', size5 => 5, size8 => 5},

  {id => 0x03, name => 'tAdd', size5 => 1, size8 => 1},
  {id => 0x04, name => 'tSub', size5 => 1, size8 => 1},
  {id => 0x05, name => 'tMul', size5 => 1, size8 => 1},
  {id => 0x06, name => 'tDiv', size5 => 1, size8 => 1},

  {id => 0x07, name => 'tPower',  size5 => 1, size8 => 1},
  {id => 0x08, name => 'tConcat', size5 => 1, size8 => 1},

  {id => 0x09, name => 'tLT', size5 => 1, size8 => 1},
  {id => 0x0A, name => 'tLE', size5 => 1, size8 => 1},
  {id => 0x0B, name => 'tEQ', size5 => 1, size8 => 1},
  {id => 0x0C, name => 'tGE', size5 => 1, size8 => 1},
  {id => 0x0D, name => 'tGT', size5 => 1, size8 => 1},
  {id => 0x0E, name => 'tNE', size5 => 1, size8 => 1},

  {id => 0x0F, name => 'tIsect', size5 => 1, size8 => 1},
  {id => 0x10, name => 'tList',  size5 => 1, size8 => 1},
  {id => 0x11, name => 'tRange', size5 => 1, size8 => 1},

  {id => 0x12, name => 'tUplus',  size5 => 1, size8 => 1},
  {id => 0x13, name => 'tUminus', size5 => 1, size8 => 1},

  {id => 0x14, name => 'tPercent',  size5 => 1, size8 => 1},
  {id => 0x15, name => 'tParen',    size5 => 1, size8 => 1},
  {id => 0x16, name => 'tMissArg',  size5 => 1, size8 => 1},

  {id => 0x17, name => 'tStr',  size5 => undef, size8 => undef},
  {id => 0x18, name => 'tNlr',  size5 => undef, size8 => undef},
  {id => 0x19, name => 'tAttr', size5 => undef, size8 => undef},

  # 0x1A tSheet unused in BIFF5/8
  # 0x1B tEndSheet unused in BIFF5/8
 
  {id => 0x1C, name => 'tErr',  size5 => 2, size8 => 2},
  {id => 0x1D, name => 'tBool', size5 => 2, size8 => 2},

  {id => 0x1E, name => 'tInt',  size5 => 3, size8 => 3},
  {id => 0x1F, name => 'tNum',  size5 => 9, size8 => 9},
];

my $classifiedTokens =
[
  {id => 0x01, name => 'tFunc',    size5 => 3, size8 => 3},
  {id => 0x02, name => 'tFuncVar', size5 => 4, size8 => 4},
  {id => 0x03, name => 'tName',    size5 => 15, size8 => 5},
  {id => 0x05, name => 'tArea',    size5 => 7, size8 => 9},
  {id => 0x0C, name => 'tRefN',    size5 => 4, size8 => 5},
  {id => 0x0D, name => 'tAreaN',   size5 => 7, size8 => 9},
];

sub getFunctionById($)
{
  my ($id) = @_;

  CORE::state $functionsById = undef;

  # Build index to functions by id if needed
  if (!defined($functionsById))
  {
    $functionsById = {};
    for my $f (@{$functions})
    {
      $functionsById->{$f->{id}} = $f;
    }
  }

  return $functionsById->{$id};
}

sub tokenInfo($)
{
  my ($type) = @_;

  my $class = ($type & TOKEN_MASK_CLASS) >> TOKEN_SHIFT_CLASS;
  my $id    = ($type & TOKEN_MASK_ID);

  my $list = ($class == 0) ? $basicTokens : $classifiedTokens;
  for my $t (@{$list})
  {
    ($t->{id} == $id) && return $t;
  }

  return undef;
}

sub tokenTypeToClass($)
{
  my ($type) = @_;

  my $class = ($type & TOKEN_MASK_CLASS) >> TOKEN_SHIFT_CLASS;
  return $class;
}

sub tokenTypeToId($)
{
  my ($type) = @_;

  my $id    = ($type & TOKEN_MASK_ID);
  return $id;
}

sub dumpToken($)
{
  my ($token) = @_;

  my $class = ($token->{type} & TOKEN_MASK_CLASS) >> TOKEN_SHIFT_CLASS;
  my $id    = ($token->{type} & TOKEN_MASK_ID);

  my $info = tokenInfo($token->{type});
  my $name = defined($info) ? $info->{name} : '?';

  printf("Token 0x%02X class %d id 0x%02X name '%s'\n", $token->{type}, $class, $id, $name);

  return;
}

sub decodeTokens($$)
{
  my ($bytes, $version) = @_;

  my $decoder = StructDecoder->new($bytes);

  my $tokens = [];

  while ($decoder->bytesLeft())
  {
    my $type = $decoder->decodeField('u8');
    my $class = ($type & TOKEN_MASK_CLASS) >> TOKEN_SHIFT_CLASS;
    my $id    = ($type & TOKEN_MASK_ID);

    my $token = {type => $type, class => $class, id => $id};

    my $info = tokenInfo($type);
    (!defined($info)) && die(sprintf("Unknown token 0x%02X", $type));

    #printf("Token 0x%02X class %d id 0x%02X name '%s'\n", $type, $class, $id, $info->{name});

    if ($class == 0)
    {
      # Basic tokens
      if (($id >= 0x03) && ($id <= 0x16))
      {
        # Trivial single-byte tokens
      }
      elsif ($id == 0x17)
      {
        # tStr
        if ($version == Spreadsheet::Nifty::XLS::BOF_VERSION_BIFF5)
        {
          $token->{str} = Spreadsheet::Nifty::XLS::Decode::decodeLen8AnsiString($decoder);
        }
        elsif ($version == Spreadsheet::Nifty::XLS::BOF_VERSION_BIFF8)
        {
          $token->{str} = Spreadsheet::Nifty::XLS::Decode::decodeShortXLUnicodeString($decoder);
        }
        else
        {
          ...;
        }
      }
      elsif ($id == 0x19)
      {
        # tAttr - Special attributes, meaning depends on flags byte
        my $flags = $decoder->decodeField('u8');
        $token->{attr}->{flags} = $flags;
        #printf("  flags 0x%02X\n", $flags);

        # Many of these can be ignored for our purposes, but we still need to
        #  account for each flags possiblity to know how many bytes to read.
        if ($flags == 0x01)
        {
          # tAttrVolatile
          #printf("  tAttrVolatile\n");
          $decoder->decodeField('u16');  # Unused
        }
        elsif ($flags == 0x02)
        {
          # tAttrIf
          #printf("  tAttrIf\n");
          $token->{attr}->{skipCount} = $decoder->decodeField('u16');
        }
        elsif ($flags == 0x04)
        {
          # tAttrChoose
          #printf("  tAttrChoose\n");
          $token->{attr}->{choose} = $decoder->decodeHash(['choiceCount:u16', 'jumpTable:u16[choiceCount]', 'skipCount:u16']);
        }
        elsif ($flags == 0x08)
        {
          # tAttrSkip
          #printf("  tAttrSkip\n");
          $token->{attr}->{skipCount} = $decoder->decodeField('u16');
        }
        elsif ($flags == 0x10)
        {
          # tAttrSum - SUM() with one parameter
          $decoder->getBytes(2);  # Skip padding
        }
        elsif ($flags == 0x020)
        {
          # tAttrAssign - Assignment within macro sheet?
          $decoder->getBytes(2);  # Skip padding
        }
        elsif (($flags == 0x040) || ($flags == 0x41))
        {
          # tAttrSpace & tAttrSpaceVolatile - Whitespace
          #printf("  tAttrSpace or tAttrSpaceVolatile\n");
          $token->{attr}->{space} = $decoder->decodeHash(['type:u8', 'count:u8']);
        }
        else
        {
          ...;
        }
      }
      elsif ($id == 0x1D)
      {
        # tBool
        $token->{value} = $decoder->decodeField('u8');
      }
      elsif ($id == 0x1E)
      {
        # tInt
        $token->{value} = $decoder->decodeField('u16');
      }
      elsif ($id == 0x1F)
      {
        # tNum
        $token->{value} = $decoder->decodeField('f64');
      }
      else
      {
        ...;
      }
    }
    else
    {
      # Classified tokens
      if ($id == 0x01)
      {
        # tFunc
        $token->{func} = $decoder->decodeHash(['func:u16']);
      }
      elsif ($id == 0x02)
      {
        # tFuncVar
        $token->{func} = $decoder->decodeHash(['argCount:u8', 'func:u16']);
      }
      elsif ($id == 0x03)
      {
        # tName
        if ($version == Spreadsheet::Nifty::XLS::BOF_VERSION_BIFF5)
        {
          $token->{index} = $decoder->decodeField('u16');
          $decoder->getBytes(12);  # Trailing padding
        }
        elsif ($version == Spreadsheet::Nifty::XLS::BOF_VERSION_BIFF8)
        {
          $token->{index} = $decoder->decodeField('u16');
          $decoder->getBytes(2);  # Trailing padding
        }
        else
        {
          ...;
        }
      }
      elsif ($id == 0x0C)
      {
        # tRefN - Relative reference to single cell on same sheet
        if ($version == Spreadsheet::Nifty::XLS::BOF_VERSION_BIFF5)
        {
          my $row = $decoder->decodeField('u16');
          my $col = $decoder->decodeField('u8');
          $token->{ref}->{row} = $row & 0x3FFF;
          $token->{ref}->{col} = $col;
          $token->{ref}->{relRow} = ($row >> 15) & 0x1;
          $token->{ref}->{relCol} = ($row >> 14) & 0x1;
        }
        elsif ($version == Spreadsheet::Nifty::XLS::BOF_VERSION_BIFF8)
        {
          my $row = $decoder->decodeField('u16');
          my $col = $decoder->decodeField('u16');
          $token->{ref}->{row} = $row;
          $token->{ref}->{col} = $col & 0xFF;
          $token->{ref}->{relRow} = ($col >> 15) & 0x1;
          $token->{ref}->{relCol} = ($col >> 14) & 0x1;
        }
        else
        {
          ...;
        }
      }
      elsif (($id == 0x05) || ($id == 0x0D))
      {
        # tArea or tAreaN - Relative reference to cell range on same sheet
        if ($version == Spreadsheet::Nifty::XLS::BOF_VERSION_BIFF5)
        {
          my $minRow = $decoder->decodeField('u16');
          my $maxRow = $decoder->decodeField('u16');
          my $minCol = $decoder->decodeField('u8');
          my $maxCol = $decoder->decodeField('u8');
          $token->{ref}->{minRow} = $minRow & 0x3FFF;
          $token->{ref}->{maxRow} = $maxRow & 0x3FFF;
          $token->{ref}->{minCol} = $minCol;
          $token->{ref}->{maxCol} = $maxCol;
          $token->{ref}->{relMinRow} = ($minRow >> 15) & 0x01;
          $token->{ref}->{relMinCol} = ($minRow >> 14) & 0x01;
          $token->{ref}->{relMaxRow} = ($maxRow >> 15) & 0x01;
          $token->{ref}->{relMaxCol} = ($maxRow >> 14) & 0x01;
        }
        elsif ($version == Spreadsheet::Nifty::XLS::BOF_VERSION_BIFF8)
        {
          my $minRow = $decoder->decodeField('u16');
          my $maxRow = $decoder->decodeField('u16');
          my $minCol = $decoder->decodeField('u16');
          my $maxCol = $decoder->decodeField('u16');
          $token->{ref}->{minRow} = $minRow;
          $token->{ref}->{maxRow} = $maxRow;
          $token->{ref}->{minCol} = $minCol & 0xFF;
          $token->{ref}->{maxCol} = $maxCol & 0xFF;
          $token->{ref}->{relMinRow} = ($minCol >> 15) & 0x01;
          $token->{ref}->{relMinCol} = ($minCol >> 14) & 0x01;
          $token->{ref}->{relMaxRow} = ($maxCol >> 15) & 0x01;
          $token->{ref}->{relMaxCol} = ($maxCol >> 14) & 0x01;
        }
        else
        {
          ...;
        }
      }
      else
      {
        ...;
      }
    }

    push(@{$tokens}, $token);
  }

  #print main::Dumper($tokens);
  #printf("UNPARSED: %s\n", unparseTokens({row => 3, col => 6}, $tokens));

  return $tokens;
}

sub unparseTokens($$)
{
  my ($context, $tokens) = @_;

  my $stack = [];
  
  for my $t (@{$tokens})
  {
    #dumpToken($t);

    if ($t->{class} == 0)
    {
      # Basic token
      if (($t->{id} >= 0x03) && ($t->{id} <= 0x11))
      {
        # Binary operators
        CORE::state $ops = ['+', '-', '*', '/', '^', '&', '<', '<=', '=', '>=', '>', '<>', ' ', ',', ':'];
        my $args = [ splice(@{$stack}, -2) ];
        my $str = $args->[0] . $ops->[$t->{id} - 0x03] . $args->[1];
        push(@{$stack}, $str);
      }
      elsif (($t->{id} >= 0x12) && ($t->{id} <= 0x13))
      {
        # Prefix unary operators
        CORE::state $ops = ['+', '-'];
        my $arg = pop(@{$stack});
        my $str = $ops->[$t->{id} - 0x12] . $arg;
        push(@{$stack}, $str);
      }
      elsif ($t->{id} == 0x14)
      {
        # tPercent - Postfix unary operator
        my $arg = pop(@{$stack});
        push(@{$stack}, "${arg}%");
      }
      elsif ($t->{id} == 0x15)
      {
        # tParen - Enclosing parentheses
        my $arg = pop(@{$stack});
        push(@{$stack}, "(${arg})");
      }
      elsif ($t->{id} == 0x16)
      {
        # tMissArg - Missing argument
        push(@{$stack}, '');
      }
      elsif ($t->{id} == 0x17)
      {
        # tStr - String literal
        my $str = "\"" . ($t->{str} =~ s/"/""/r) . "\"";
        push(@{$stack}, $str);
      }
      elsif ($t->{id} == 0x19)
      {
        # tAttr - Special attributes
        # TODO: tAttrSpace
        if ($t->{attr}->{flags} == 0x10)
        {
          # tAttrSum - Special-case encoding for SUM() with one argument
          my $arg = pop(@{$stack});
          push(@{$stack}, "SUM(${arg})");
        }
      }
      elsif ($t->{id} == 0x1D)
      {
        # tBool - Boolean literal
        push(@{$stack}, $t->{value} ? 'TRUE' : 'FALSE');
      }
      elsif (($t->{id} == 0x1E) || ($t->{id} == 0x1F))
      {
        # tInt - 16-bit unsigned integer literal
        # tNum - 64-bit IEEE float literal
        push(@{$stack}, "" . $t->{value});
      }
      else
      {
        die(sprintf("Unhandled basic token 0x%02X", $t->{id}));
      }
    }
    else
    {
      # Classified token
      if ($t->{id} == 0x01)
      {
        # tFunc - Function call with fixed number of args
        my $function = getFunctionById($t->{func}->{func});
        (!defined($function)) && die(sprintf("Unknown function 0x%02X", $t->{func}->{func}));

        (($function->{minArgs} < 0) || ($function->{minArgs} != $function->{maxArgs})) && die(sprintf("Function '%s' encoded without arg count", $function->{name}));
        my $argCount = $function->{minArgs};

        my $args = [ splice(@{$stack}, -$argCount) ];

        my $str = $function->{name} . '(' . join(',', @{$args}) . ')';
        push(@{$stack}, $str);
      }
      elsif ($t->{id} == 0x02)
      {
        # tFuncVar - Function call with variable number of args
        if ($t->{func}->{func} == 0xFF)
        {
          # Special case for user-defined or future function? First arg is tName with name of the function.
          my $args = [ splice(@{$stack}, -$t->{func}->{argCount}) ];
          my $name = shift(@{$args});
          $name =~ s#^_xlfn\.##;  # Remove future function prefix
          my $str = $name . '(' . join(',', @{$args}) . ')';
          push(@{$stack}, $str);
        }
        else
        {
          # Built-in function with variable number of args
          my $function = getFunctionById($t->{func}->{func});
          (!defined($function)) && die(sprintf("Unknown function 0x%02X", $t->{func}->{func}));
          # TODO: Check arg count vs function maxArgs

          my $args = ($t->{func}->{argCount}) ? [ splice(@{$stack}, -$t->{func}->{argCount}) ] : [];
          my $str = $function->{name} . '(' . join(',', @{$args}) . ')';
          push(@{$stack}, $str);
        }
      }
      elsif ($t->{id} == 0x03)
      {
        # tName
        #printf("tName index %d\n", $t->{index});
        (!defined($context->{sheet})) && die("No sheet context while unparsing tName token");
        my $label = $context->{sheet}->{workbook}->getLabel($t->{index} - 1);
        push(@{$stack}, $label->{name});
      }
      elsif ($t->{id} == 0x0C)
      {
        # tRefN - Relative 2D cell reference
        my $col = $t->{ref}->{relCol} ? ($context->{col} + $t->{ref}->{col}) % Spreadsheet::Nifty::XLS::MAX_COL_COUNT : $t->{ref}->{col};
        my $row = $t->{ref}->{relRow} ? ($context->{row} + $t->{ref}->{row}) % Spreadsheet::Nifty::XLS::MAX_ROW_COUNT : $t->{ref}->{row};
        #print main::Dumper($context->{col}, $context->{row}, $t, $col, $row);

        my $str = ((!$t->{ref}->{relCol}) ? '$' : '') . Spreadsheet::Nifty::Utils->colIndexToString($col) . ((!$t->{ref}->{relRow}) ? '$' : '') . int($row + 1);
        push(@{$stack}, $str);
      }
      elsif (($t->{id} == 0x05) || ($t->{id} == 0x0D))
      {
        # tArea or tAreaN - Relative 2D area reference
        my $minCol = $t->{ref}->{relMinCol} ? ($context->{col} + $t->{ref}->{minCol}) % Spreadsheet::Nifty::XLS::MAX_COL_COUNT : $t->{ref}->{minCol};
        my $maxCol = $t->{ref}->{relMaxCol} ? ($context->{col} + $t->{ref}->{maxCol}) % Spreadsheet::Nifty::XLS::MAX_COL_COUNT : $t->{ref}->{maxCol};
        my $minRow = $t->{ref}->{relMinRow} ? ($context->{row} + $t->{ref}->{minRow}) % Spreadsheet::Nifty::XLS::MAX_ROW_COUNT : $t->{ref}->{minRow};
        my $maxRow = $t->{ref}->{relMaxRow} ? ($context->{row} + $t->{ref}->{maxRow}) % Spreadsheet::Nifty::XLS::MAX_ROW_COUNT : $t->{ref}->{maxRow};
        #print main::Dumper($context->{col}, $context->{row}, $t, $minCol, $minRow, $maxCol, $maxRow);
        my $str = sprintf("%s%s%s%s:%s%s%s%s",
         $t->{ref}->{relMinCol} ? '' : '$',
         Spreadsheet::Nifty::Utils->colIndexToString($minCol),
         $t->{ref}->{relMinRow} ? '' : '$',
         int($minRow + 1),
         $t->{ref}->{relMaxCol} ? '' : '$',
         Spreadsheet::Nifty::Utils->colIndexToString($maxCol),
         $t->{ref}->{relMaxRow} ? '' : '$',
         int($maxRow + 1));
        push(@{$stack}, $str);
      }
      else
      {
        die(sprintf("Unhandled classified token 0x%02X", $t->{id}));
      }
    }

    #print main::Dumper('STACK', $stack);
  }

  my $stackLength = scalar(@{$stack});
  ($stackLength != 1) && die("unparseTokens(): Expected stack length 1, got $stackLength");

  return $stack->[0];
}

1;
