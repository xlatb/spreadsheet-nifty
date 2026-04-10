#!/usr/bin/perl -w
use warnings;
use strict;

package Spreadsheet::Nifty::XLSX::Cell;
use parent 'Spreadsheet::Nifty::Cell';

# === Class methods ===

sub new($$$;$)
{
  my $class = shift();
  my ($type, $value, $private) = @_;

  my $self = $class->SUPER::new($type, $value);
  $self->{p} = $private // {};

  return $self;
}

# === Instance methods ===

sub getFormula()
{
  my $self = shift();

  return $self->{p}->{f};
}

sub getFgColor()
{
  my $self = shift();

  return $self->getFillColor('fgColor');
}

sub getBgColor()
{
  my $self = shift();

  return $self->getFillColor('bgColor');
}

sub getFillColor($)
{
  my $self = shift();
  my ($name) = @_;

  (!defined($self->{p}->{xf})) && return undef;  # No direct style

  my $xf = $self->{p}->{ctx}->workbook()->getXf($self->{p}->{xf});
  (!defined($xf->{fillId})) && return undef;  # No fill info

  my $fill = $self->{p}->{ctx}->workbook()->getFill($xf->{fillId});
  (!defined($fill->{bgColor})) && return undef;  # No fill background colour

  return $self->{p}->{ctx}->workbook()->resolveColor($fill->{bgColor});
}

1;
