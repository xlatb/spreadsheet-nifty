#!/usr/bin/perl -w
use warnings;
use strict;

use Scalar::Util qw();

package Spreadsheet::Nifty::XLSX::Cell;
use parent 'Spreadsheet::Nifty::Cell';

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

  my $xf = $self->{ctx}->()->workbook()->getXf($self->{p}->{xf});
  (!defined($xf->{fillId})) && return undef;  # No fill info

  my $fill = $self->{ctx}->()->workbook()->getFill($xf->{fillId});
  (!defined($fill->{bgColor})) && return undef;  # No fill background colour

  return $self->{ctx}->()->workbook()->resolveColor($fill->{bgColor});
}

1;
