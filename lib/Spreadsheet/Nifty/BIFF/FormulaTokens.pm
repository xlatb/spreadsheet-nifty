#!/usr/bin/perl -w
use warnings;
use strict;

package Spreadsheet::Nifty::BIFF::FormulaTokens;

use Spreadsheet::Nifty::BIFF::Constants;

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
  {id => 0x04, name => 'tRef',     size5 => 3, size8 => 4},
  {id => 0x05, name => 'tArea',    size5 => 7, size8 => 9},
  {id => 0x0C, name => 'tRefN',    size5 => 4, size8 => 5},
  {id => 0x0D, name => 'tAreaN',   size5 => 7, size8 => 9},
  {id => 0x1A, name => 'tRef3d',   size5 => 17, size8 => 6},
  {id => 0x1B, name => 'tArea3d',  size5 => 20, size8 => 10},
];

# === Class methods ===

sub new($)
{
  my $class = shift();
  my ($version) = @_;

  my $self = {};
  $self->{version} = $version;

  if ($version == Spreadsheet::Nifty::BIFF::Constants::BIFF_V5)
  {
    $self->{maxColCount} = 256;    # 8 bits
    $self->{maxRowCount} = 16384;  # 14 bits
  }
  elsif ($version == Spreadsheet::Nifty::BIFF::Constants::BIFF_V8)
  {
    $self->{maxColCount} = 256;    # 8 bits
    $self->{maxRowCount} = 65536;  # 14 bits
  }
  elsif ($version == Spreadsheet::Nifty::BIFF::Constants::BIFF_V12)
  {
    $self->{maxColCount} = 16384;    # 16 bits
    $self->{maxRowCount} = 1048576;  # 20 bits
  }
  else
  {
    die("Unknown BIFF version");
  }

  bless($self, $class);
  return $self;
}

# === Instance methods ===

sub tokenInfo($)
{
  my ($type) = @_;

  my $class = ($type & TOKEN_MASK_CLASS) >> TOKEN_SHIFT_CLASS;
  my $id    = ($type & TOKEN_MASK_ID);

  my $list = ($class == TOKEN_CLASS_BASIC) ? $basicTokens : $classifiedTokens;
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

  my $id = ($type & TOKEN_MASK_ID);
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

1;
