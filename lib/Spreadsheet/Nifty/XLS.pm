#!/usr/bin/perl -w
use warnings;
use strict;

package Spreadsheet::Nifty::XLS;

use constant
{
  BOF_TYPE_WORKBOOK  => 0x05,
  BOF_TYPE_WORKSHEET => 0x10,  # NOTE: Also used for dialog sheets
  BOF_TYPE_CHART     => 0x20,
  BOF_TYPE_MACRO     => 0x40,
};

use constant
{
  BOF_VERSION_BIFF2 => 0x0200,
  BOF_VERSION_BIFF3 => 0x0300,
  BOF_VERSION_BIFF4 => 0x0400,
  BOF_VERSION_BIFF5 => 0x0500,
  BOF_VERSION_BIFF8 => 0x0600,  # Yes, BIFF8 is stored as 6, not 8
};

use constant
{
  ROW_BLOCK_SIZE  => 32,
  ROW_BLOCK_SHIFT => 5,  # == log2(ROW_BLOCK_SIZE)
};

use constant
{
  MAX_COL_COUNT => 256,
  MAX_ROW_COUNT => 65536,
};

# Excel will silently try this default password.
our $defaultPassword = 'VelvetSweatshop';

1;
