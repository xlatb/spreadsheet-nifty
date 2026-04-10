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

sub formula()
{
  my $self = shift();

  return $self->{p}->{f};
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

sub fgColor()
{
  my $self = shift();

  return $self->getFillColor('fgColor');
}

sub bgColor()
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
  (!defined($xf->{fillId})) && return undef;  # No fill applied

  my $fill = $self->{ctx}->()->workbook()->getFill($xf->{fillId});
  (!defined($fill)) && return undef;  # No such fill

  my $color = $fill->{$name};
  (!defined($color)) && return undef;  # No colour by that name

  return $self->{ctx}->()->workbook()->resolveColor($color);
}

1;
