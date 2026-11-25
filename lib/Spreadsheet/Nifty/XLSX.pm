#!/usr/bin/perl -w
use warnings;
use strict;

package Spreadsheet::Nifty::XLSX;

our $namespaces = 
{
  main             => 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',

  xm               => 'http://schemas.microsoft.com/office/excel/2006/main',
  x14              => 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main',  # Extensions from Excel 2010
  x15              => 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main',  # Extensions from Excel 2013
};

# NOTE: These can be published with mixed case. Make sure to uppercase before
#  comparison.
our $extensionIds =
{
  'x14:dataValidations'        => '{CCE6A557-97BC-4B89-ADB6-D9C93CAAB3DF}',
  'x14:protectedRanges'        => '{FC87AEE6-9EDD-498A-B2FB-16103846CE08}',
  'x14:conditionalFormattings' => '{78C0D931-6437-407D-A8EE-F0AAD7539E65}',
  'x14:sparklineGroups'        => '{05C60535-1F16-4FD2-B633-F4F36F0B64E0}',
  'x14:ignoredErrors'          => '{01252117-D84E-4E92-8308-4BE1C098FCBB}',
  'x14:workbookPr'             => '{7523E5D3-25F3-A5E0-1632-64F254C22452}',

  'x15:workbookPr'             => '{FF2C5DFA-F100-4F1F-AC45-02E685F130E8}',
  'x16r2:dynamicArrayPr'       => '{20071BC0-6874-4D63-882D-4C244E6B4F43}',
};

1;
