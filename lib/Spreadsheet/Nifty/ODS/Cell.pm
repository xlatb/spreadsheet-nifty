#!/usr/bin/perl -w
use warnings;
use strict;

package Spreadsheet::Nifty::ODS::Cell;
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

sub stringValue()
{
  my $self = shift();

  if ($self->{t} == Spreadsheet::Nifty::TYPE_DATE)
  {
    if (defined($self->{p}->{value}))
    {
      return $self->{p}->{value};
    }
  }

  return $self->SUPER::stringValue();
}

# TODO: We need to tokenize and rewrite the formula
sub formula()
{
  return undef;
}

# TODO: We need to unparse the format string
sub formatString()
{
  return undef;
}

sub fgColor()
{
  my $self = shift();

  (!defined($self->{p}->{style})) && return undef;  # Unstyled cell

  my $style = $self->{ctx}->()->{fileReader}->getStyle('table-cell', $self->{p}->{style});

  if (defined($style) && defined($style->{text}) && defined($style->{text}->{color}))
  {
    my $fgColor = $style->{text}->{color};
    if ($fgColor =~ m/^#([0-9A-F]{6})$/i)
    {
      return uc($1) . 'FF';  # RRGGBBAA
    }
  }

  return undef;
}

sub bgColor()
{
  my $self = shift();

  (!defined($self->{p}->{style})) && return undef;  # Unstyled cell

  my $style = $self->{ctx}->()->{fileReader}->getStyle('table-cell', $self->{p}->{style});

  if (defined($style) && defined($style->{cell}) && defined($style->{cell}->{'background-color'}))
  {
    my $bgColor = $style->{cell}->{'background-color'};
    if ($bgColor =~ m/^#([0-9A-F]{6})$/i)
    {
      return uc($1) . 'FF';  # RRGGBBAA
    }
  }

  return undef;
}

1;
