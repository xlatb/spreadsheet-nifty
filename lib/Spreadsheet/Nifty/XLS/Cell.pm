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

sub new($$$;$)
{
  my $class = shift();
  my ($type, $value, $ctx, $private) = @_;

  my $self = $class->SUPER::new($type, $value);
  $self->{p} = $private // {};
  $self->{ctx} = sub { return $ctx; };  # Closure around context

  Scalar::Util::weaken($ctx);

  return $self;
}

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

sub formatString()
{
  my $self = shift();

  (!defined($self->{p}->{xf})) && return undef;  # No direct style

  my $xf = $self->{ctx}->()->workbook()->getXf($self->{p}->{xf});
  (!defined($xf->{numberFormatId})) && return undef;  # No number format applied

  my $numberFormat = $self->{ctx}->()->workbook()->getNumberFormat($xf->{numberFormatId});
  return $numberFormat;
}

sub formula()
{
  my $self = shift();

  (!defined($self->{p}->{formula})) && return undef;

  my $formulaDecoder = $self->{ctx}->()->getFormulaDecoder();
  my $formula = $formulaDecoder->decode($self->{p}->{formula}->{tokenStream}, $self->{p}->{formula}->{extra});
  #print main::Dumper($self, $formula);
  my $context = {sheetReader => $self->{ctx}->(), row => $self->{p}->{row}, col => $self->{p}->{col}};

  return $formula->unparse($context);
}

1;
