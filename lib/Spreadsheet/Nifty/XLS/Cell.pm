#!/usr/bin/perl -w
use warnings;
use strict;

package Spreadsheet::Nifty::XLS::Cell;
use parent 'Spreadsheet::Nifty::Cell';

my $errorMap =
{
   0 => 1,  # #NULL!
   7 => 2,  # #DIV/0!
  15 => 3,  # #VALUE!
  23 => 4,  # #REF!
  29 => 5,  # #NAME?
  36 => 6,  # #NUM!
  42 => 7,  # #NA!
};

# === Class methods ===

# === Instance methods ===

sub value()
{
  my $self = shift();

  if ($self->{t} == Spreadsheet::Nifty::TYPE_ERR)
  {
    return $errorMap->{$self->{v}};
  }

  return $self->{v};
}

1;
