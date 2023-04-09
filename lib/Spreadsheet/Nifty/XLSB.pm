#!/usr/bin/perl -w
use warnings;
use strict;

package Spreadsheet::Nifty::XLSB;

our $namespaces =
{
  main             => 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  sharedStrings    => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',
};

1;
