#!/usr/bin/perl -w
use warnings;
use strict;

# Utility functions for use with XML::LibXML::Reader.
package Spreadsheet::Nifty::XMLReaderUtils;

use XML::LibXML qw(:libxml);
use XML::LibXML::Reader;

# === Class methods ===

sub dump($;$)
{
  my $class = shift();
  my ($xmlReader, $prefix) = @_;

  if (defined($prefix))
  {
    printf("%s: ", $prefix);
  }

  printf("depth %d empty %d nodeType %d localname %s\n", $xmlReader->depth(), $xmlReader->isEmptyElement(), $xmlReader->nodeType(), $xmlReader->localName());
}

# Given an XML Reader, if we're positioned at the beginning of an element, jumps to the end of the element.
sub skipElement($)
{
  my $class = shift();
  my ($xmlReader) = @_;

  ($xmlReader->isEmptyElement()) && return;  # Empty element has no content to skip

  my $nodeType = $xmlReader->nodeType();
  ($nodeType == XML_READER_TYPE_END_ELEMENT) && return;  # Already at the end of the element

  # If we're not even pointing at an element, give up
  ($nodeType != XML_READER_TYPE_ELEMENT) && die("skipElement(): Expected an element node, but got type ${nodeType}");

  my $depth = $xmlReader->depth();

  my $state;
  while (($state = $xmlReader->read()) == 1)
  {
    #printf("skipElement() read: depth %d empty %d nodeType %d localname %s\n", $xmlReader->depth(), $xmlReader->isEmptyElement(), $xmlReader->nodeType(), $xmlReader->localName());
    ($xmlReader->depth() == $depth) && ($xmlReader->nodeType() == XML_READER_TYPE_END_ELEMENT) && return;  # Found the end

    $xmlReader->skipSiblings();
    #printf("skipElement() skipSiblings: depth %d empty %d nodeType %d localname %s\n", $xmlReader->depth(), $xmlReader->isEmptyElement(), $xmlReader->nodeType(), $xmlReader->localName());
    ($xmlReader->depth() == $depth) && ($xmlReader->nodeType() == XML_READER_TYPE_END_ELEMENT) && return;  # Found the end
  }

  die("skipElement(): Read error");
}

# Finds the first element with the given name and/or namespace which is a
#  child of the current element.
# Returns true if such an element was found, and false otherwise.
# Will not advance past the end of the current element.
sub findChildElement($;$$)
{
  my $class = shift();
  my ($xmlReader, $name, $ns) = @_;

  ($xmlReader->isEmptyElement()) && return !!0;  # Empty element has no children

  # If we're not even pointing at an element, give up
  my $nodeType = $xmlReader->nodeType();
  ($nodeType != XML_READER_TYPE_ELEMENT) && die("findChildElement(): Expected an element node, but got type ${nodeType}");

  my $depth = $xmlReader->depth();

  # Try to step into first child node
  my $state  = $xmlReader->read();
  ($xmlReader->depth() == $depth) && return !!0;  # Likely immediate end tag

  # Is the current node a match?
  if (($xmlReader->nodeType() == XML_READER_TYPE_ELEMENT) &&
      (!defined($name) || ($name eq $xmlReader->localName())) &&
      (!defined($ns) || ($ns eq $xmlReader->namespaceURI())))
  {
    return !!1;  # Found
  }

  # It wasn't the first child, look for a sibling
  return $class->findSiblingElement($xmlReader, $name, $ns);
}

# Finds the next element with the given name and/or namespace which is a
#  sibling of the current node.
# Returns true if such an element was found, and false otherwise.
# Will not advance past the end of the parent element.
# NOTE: Why not use XML::LibXML::Reader's nextSiblingElement method?
#  Unfortunately it advances an extra step in the case of no match.
sub findSiblingElement($;$$)
{
  my $class = shift();
  my ($xmlReader, $name, $ns) = @_;

  my $depth = $xmlReader->depth();

  my $state;
  while (($state = $xmlReader->read()) == 1)
  {
    #printf("findSiblingElement(): depth %d empty %d nodeType %d localname %s\n", $xmlReader->depth(), $xmlReader->isEmptyElement(), $xmlReader->nodeType(), $xmlReader->localName());
    my $nodeType = $xmlReader->nodeType();
    if (($nodeType == XML_READER_TYPE_ELEMENT) &&
        (!defined($name) || ($name eq $xmlReader->localName())) &&
        (!defined($ns) || ($ns eq $xmlReader->namespaceURI())))
    {
      return !!1;  # Found a match
    }
    elsif (($nodeType == XML_READER_TYPE_END_ELEMENT) && ($xmlReader->depth() < $depth))
    {
      return !!0;  # Reached parent end tag
    }

    if (($nodeType == XML_READER_TYPE_ELEMENT) && (!$xmlReader->isEmptyElement()))
    {
      $xmlReader->read();
      $xmlReader->skipSiblings();
    }
  }

  die("findSiblingElement(): Read error");
}

sub atStartOfElement($$$)
{
  my $class = shift();
  my ($xmlReader, $name, $ns) = @_;

  ($xmlReader->nodeType() != XML_READER_TYPE_ELEMENT) && return !!0;
  (defined($name) && ($name ne $xmlReader->localName())) && return !!0;
  (defined($ns) && ($ns ne $xmlReader->namespaceURI())) && return !!0;

  return !!1;
}

# Returns true if we're positioned at the end of an element matching the given
#  name and/or namespace.
sub atEndOfElement($$$)
{
  my $class = shift();
  my ($xmlReader, $name, $ns) = @_;

  my $nodeType = $xmlReader->nodeType();
  if ($nodeType == XML_READER_TYPE_ELEMENT)
  {
    (!$xmlReader->isEmptyElement()) && return !!0;
  }
  elsif ($nodeType != XML_READER_TYPE_END_ELEMENT)
  {
    return !!0;
  }

  (defined($name) && ($name ne $xmlReader->localName())) && return !!0;
  (defined($ns) && ($ns ne $xmlReader->namespaceURI())) && return !!0;

  return !!1;
}

# Skip content until we reach the given depth and returns true.
# If we are already at the specified depth or less, returns false and nothing happens.
sub ascendToDepth($$)
{
  my $class = shift();
  my ($xmlReader, $depth) = @_;

  ...;  # TODO: Untested
  ($xmlReader->depth() <= $depth) && return !!0;

  my $state;
  while (($state = $xmlReader->read()) == 1)
  {
    printf("ascendToDepth() loop: depth %d empty %d nodeType %d localname %s\n", $xmlReader->depth(), $xmlReader->isEmptyElement(), $xmlReader->nodeType(), $xmlReader->localName());
    my $nodeType = $xmlReader->nodeType();

    ($xmlReader->depth() == $depth) && ($xmlReader->nodeType() == XML_READER_TYPE_END_ELEMENT) && return !!1;  # Found

    $xmlReader->skipSiblings();
    printf("ascendToDepth() skipSiblings: depth %d empty %d nodeType %d localname %s\n", $xmlReader->depth(), $xmlReader->isEmptyElement(), $xmlReader->nodeType(), $xmlReader->localName());
  }

  die("ascendToDepth(): Read error");
}

# Skip content until we reach an ancestor element's end tag matching the given name and/or namespace.
sub ascendToElement($$$)
{
  my $class = shift();
  my ($xmlReader, $name, $ns) = @_;

  my $state;
  while (($state = $xmlReader->skipSiblings()) == 1)
  {
    #printf("ascendToElement() loop: depth %d empty %d nodeType %d localname %s\n", $xmlReader->depth(), $xmlReader->isEmptyElement(), $xmlReader->nodeType(), $xmlReader->localName());

    # NOTE: We don't consider a matching empty element to be an end condition
    #  because we can't have descended into it.
    my $nodeType = $xmlReader->nodeType();
    if (($nodeType == XML_READER_TYPE_END_ELEMENT) &&
        (!defined($name) || ($name eq $xmlReader->localName())) &&
        (!defined($ns) || ($ns eq $xmlReader->namespaceURI())))
    {
      return !!1;  # Found
    }
  }

  die("ascendToElement(): Read error");
}

1;
