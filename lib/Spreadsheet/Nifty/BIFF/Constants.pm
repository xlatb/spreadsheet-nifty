#!/usr/bin/perl -w
use warnings;
use strict;

package Spreadsheet::Nifty::BIFF::Constants;

# Locally-defined, not standard.
use constant
{
  BIFF_V5  => 5,   # XLS (Excel 5.0)
  BIFF_V8  => 8,   # XLS (Excel 97)
  BIFF_V12 => 12,  # XLSB
};

1;
