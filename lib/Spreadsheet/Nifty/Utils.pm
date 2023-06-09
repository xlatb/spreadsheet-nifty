#!/usr/bin/perl -w
use warnings;
use strict;

package Spreadsheet::Nifty::Utils;

use POSIX qw();

# POSIX epoch 1970-01-01 as an Excel day number, measured from different base years.
my $timeEpochs = { 1900 => 25569, 1904 => 24107 };

# Converts a POSIX timestamp (interpreted as UTC), to an Excel time value.
sub posixToExcelTime($;$)
{
  my $class = shift();
  my ($time, $baseyear) = @_;

  my $epoch = $timeEpochs->{$baseyear // '1900'};
  (!defined($epoch)) && die("Unexpected base year");

  my $excel = $epoch + ($time / (3600 * 24));
  return $excel;
}

# Converts an Excel time value to a POSIX timestamp in UTC.
sub excelToPosixTime($;$)
{
  my $class = shift();
  my ($excel, $baseyear) = @_;

  my $epoch = $timeEpochs->{$baseyear // '1900'};
  (!defined($epoch)) && die("Unexpected base year");
  
  # NOTE: int() truncates towards zero, which gives incorrect results for
  #  negative values, so we use POSIX::floor() instead.
  my $time = POSIX::floor((($excel - $epoch) * (3600 * 24)) + 0.5);
  return $time;
}

# All inputs are integers. Sign should be one for negative and zero for positive.
# For IEEE single: ieeePartsToValue(sign, exponent, mantissa, 8, 23, 127)
# For IEEE double: ieeePartsToValue(sign, exponent, mantissa, 11, 52, 1023)
sub ieeePartsToValue($$$$$$)
{
  my $class = shift();
  my ($sign, $exponent, $mantissa, $ebits, $mbits, $ebias) = @_;

  $sign = ($sign) ? -1 : 1;

  if ($exponent == 0)
  {
    # Denormal
    $exponent = 1 - $ebias;
  }
  elsif ($exponent == ((1 << $ebits) - 1))
  {
    # Special
    if ($mantissa == 0)
    {
      return (($sign == 1) ? "+inf" : "-inf") + 0.0;
    }
    else
    {
     return (($sign == 1) ? "+nan" : "-nan") + 0.0;
    }
  }
  else
  {
    # Normal
    $exponent -= $ebias;  # Unbias
    $mantissa |= (1 << $mbits);  # Add implied leading 1 bit
  }

  $mantissa /= (1 << $mbits);
  my $value =  $sign * (2 ** $exponent) * ($mantissa);

  #printf("  IEEE sign %d exponent %d mantissa %g = %g\n", $sign, $exponent, $mantissa, $value);
  return $value;
}

# An Rk is a number packed into a 32-bit value in a goofy way.
sub translateRk($)
{
  my $class = shift();
  my ($value) = @_;

  my $flagA = $value & 0x01;
  my $flagB = $value & 0x02;
  $value &= 0xFFFFFFFC;  # Discard flag bits

  #printf("  Intermediate: 0x%08X\n", $value);
  if ($flagB)
  {
    # 30-bit signed integer
    $value >>= 2;
    if ($value & 0x20000000)  # Sign bit
    {
      $value = -(($value ^ 0x3FFFFFFF) + 1);  # Flip to negative
    }
    #printf("  Signed integer: 0x%08X (%d)\n", $value, $value);
  }
  else
  {
    # This worked but is unlikely to be portable
    #my $packed = pack('C8', reverse(unpack('C8', pack('N', $value) . pack('V', 0))));
    #$value = unpack('d', $packed');
    my $sign     = ($value & 0x80000000) >> 31;
    my $exponent = ($value & 0x7FF00000) >> 20;
    my $mantissa = ($value & 0x000FFFFF) << 32;
    #printf("  parts: sign %d exponent %d mantissa %d\n", $sign, $exponent, $mantissa);
    $value = Spreadsheet::Nifty::Utils->ieeePartsToValue($sign, $exponent, $mantissa, 11, 52, 1023);
  }

  if ($flagA)
  {
    #printf("  Divide by 100\n");
    $value = $value / 100;
  }
  #printf("  Final value: %s\n", $value);

  return $value;
}

# Translates a zero-based column index to a column string.
sub colIndexToString($)
{
  my $class = shift();
  my ($index) = @_;

  (!defined($index)) && return undef;
  (($index < 0) || ($index > 16384)) && return undef;

  my $digits = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';

  my $letters = '';
  
  $index += 1;  # Bias
  
  while ($index > 0)
  {
    $letters = substr($digits, ($index % 26) - 1, 1) . $letters;
    $index = int(($index - 1) / 26);
  }

  return $letters;
}

# Translates a column string to a zero-based column index.
sub stringToColIndex($)
{
  my $class = shift();
  my ($letters) = @_;

  my $digits = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';

  $letters = uc($letters);

  my $len = length($letters);
  (($len < 1) || ($len > 3)) && return undef;  # Out of range

  my $index = 0;

  for (my $i = 0; $i < $len; $i++)
  {
    $index = $index * 26;

    my $j = index($digits, substr($letters, $i, 1));
    ($j == -1) && return undef;  # Bad character in string

    $index += ($j + 1);
  }

  return $index - 1;
}

1;
