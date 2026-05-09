#!/usr/bin/perl -w
use warnings;
use strict;

package Spreadsheet::Nifty::ODS;

our $mimetypes =
[
  'application/vnd.oasis.opendocument.spreadsheet',
  'application/vnd.oasis.opendocument.spreadsheet-template',
];

our $namespaces =
{
  'number'   => 'urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0',
  'office'   => 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
  'table'    => 'urn:oasis:names:tc:opendocument:xmlns:table:1.0',
  'text'     => 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
  'style'    => 'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
  'fo'       => 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0',
  'loext'    => 'urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0',
  'tableooo' => 'http://openoffice.org/2009/table',
};

1;
