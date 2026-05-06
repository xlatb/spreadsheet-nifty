#!/usr/bin/perl -w
use warnings;
use strict;

package Spreadsheet::Nifty::DateUtils;

# Year, month, and day are one-based.
# Returned day of year is zero-based.
sub dayOfYear($$$)
{
  my $class = shift();
  my ($year, $month, $day) = @_;

  CORE::state $monthStartDays = [0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334];

  my $doy = $monthStartDays->[$month - 1] + ($day - 1);

  if ($class->isLeapYear($year) && ($month >= 3))
  {
    $doy++;
  }

  return $doy;
}

sub isLeapYear($)
{
  my $class = shift();
  my ($year) = @_;

  if ($year % 4 == 0)
  {
    if ($year % 100 == 0)
    {
      if ($year % 400 == 0)
      {
        return 1;
      }

      return 0;
    }

    return 1;
  }

  return 0;
}

sub daysInYear($)
{
  my $class = shift();
  my ($year) = @_;

  return $class->isLeapYear($year) ? 366 : 365;
}

1;
