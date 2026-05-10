#!/usr/bin/perl -w
use warnings;
use strict;

package Spreadsheet::Nifty::XLSB::Cell;
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

# Given a BrtColor structure, returns the RGBA values for the colour.
sub resolveColor($)
{
  my $self = shift();
  my ($color) = @_;

  # If bundled RGBA is valid, just return that
  if ($color->{type} & 0x01)
  {
    return {r => $color->{red}, g => $color->{green}, b => $color->{blue}, a => $color->{alpha}};
  }

  my ($r, $g, $b, $a);
  my $type = $color->{type} >> 1;  # NOTE: Colour type 0x02 should have been handled above
  if ($type == 0x01)  # Indexed colour
  {
    my $c = $self->{ctx}->()->{workbook}->{styles}->{palette}->{indexed}->getColorRGB($color->{index});
    ($r, $g, $b, $a) = ($c->{r}, $c->{g}, $c->{b}, 255);
  }
  elsif ($type == 0x03)  # Theme colour
  {
    # TODO: Theme colours
    return undef;
  }

  # TODO: Process tint field

  return {r => $r, g => $g, b => $b, a => $a};
}

sub fgColor()
{
  my $self = shift();

  my $xf = $self->{ctx}->()->{workbook}->{styles}->getXf($self->{p}->{xf});
  (!defined($xf)) && return undef;

  my $fill = $self->{ctx}->()->{workbook}->{styles}->getFill($xf->{fillId});
  (!defined($fill)) && return undef;

  my $color = $self->resolveColor($fill->{fgColor});
  (!defined($color)) && return undef;

  return sprintf("%02X%02X%02X%02X", $color->{r}, $color->{g}, $color->{b}, $color->{a});
}

sub bgColor()
{
  my $self = shift();

  my $xf = $self->{ctx}->()->{workbook}->{styles}->getXf($self->{p}->{xf});
  (!defined($xf)) && return undef;

  my $fill = $self->{ctx}->()->{workbook}->{styles}->getFill($xf->{fillId});
  (!defined($fill)) && return undef;

  my $color = $self->resolveColor($fill->{bgColor});
  (!defined($color)) && return undef;

  return sprintf("%02X%02X%02X%02X", $color->{r}, $color->{g}, $color->{b}, $color->{a});
}

sub formatString()
{
  my $self = shift();

  return undef;  #TODO
}

sub formula()
{
  my $self = shift();

  return undef;  #TODO
}


1;
