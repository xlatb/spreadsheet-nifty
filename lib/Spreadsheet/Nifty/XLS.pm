#!/usr/bin/perl -w
use warnings;
use strict;

package Spreadsheet::Nifty::XLS;

use Spreadsheet::Nifty::BIFF::Constants;

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
  BOF_VERSION_BIFF5 => 0x0500,  # Also used by BIFF7
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

# Translate the BIFF version number as encoded in the BOF field to one of our
#  own version constants.
sub bofToBiffVersion($)
{
  my ($bof) = @_;

  # NOTE: We don't use => here because we don't want to stringify the keys
  CORE::state $versionMap =
  {
    BOF_VERSION_BIFF5 , Spreadsheet::Nifty::BIFF::Constants::BIFF_V5,
    BOF_VERSION_BIFF8 , Spreadsheet::Nifty::BIFF::Constants::BIFF_V8
  };

  return $versionMap->{$bof};
}


1;
