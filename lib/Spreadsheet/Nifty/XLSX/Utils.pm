#!/usr/bin/perl -w
use warnings;
use strict;

use Spreadsheet::Nifty::XLSX::FileReader;

package Spreadsheet::Nifty::XLSX::Utils;

sub unprotect($$)
{
  my $class = shift();
  my ($inputFile, $outputFile) = @_;

  # Open file
  my $reader = Spreadsheet::Nifty::XLSX::FileReader->new();
  (!$reader->open($inputFile)) && return 0;

  # Process each worksheet member
  my $sheetCount = $reader->getSheetCount();
  for (my $i = 0; $i < $sheetCount; $i++)
  {
    $class->unprotectWorksheetMember($reader, $i);
  }

  # Set each member's compression method
  for my $member ($reader->{zipPackage}->{zip}->members())
  {
    if ($member->fileName() =~ m#[.](rels|xml)$#i)
    {
      $member->desiredCompressionMethod(Archive::Zip::COMPRESSION_DEFLATED);
    }
  }

  # Write to temp file
  my $tempFile = $outputFile . '.part';
  my $status = $reader->{zipPackage}->{zip}->writeToFileNamed({filename => $tempFile});
  ($status != Archive::Zip::AZ_OK) && return 0;

  # Rename to final output filename
  $status = rename($tempFile, $outputFile);
  (!$status) && return 0;

  return 1;
}

sub unprotectWorksheetMember($$)
{
  my $class = shift();
  my ($fileReader, $index) = @_;

  #printf("Unprotecting worksheet %d\n", $index);

  # Find worksheet's relationship
  my $relId = $fileReader->{workbook}->{sheets}->[$index]->{id};
  my $rel = $fileReader->{workbook}->{relationships}->{$relId};
  (!defined($rel)) && return 0;

  # Open the worksheet member for reading
  my $zipReader = $fileReader->{zipPackage}->openMember($rel->{partname});
  (!$zipReader) && return 0;

  # Parse entire document
  my $xml = XML::LibXML->new({validation => 0, expand_entities => 0, no_network => 1});
  my $doc = $xml->parse_fh($zipReader);
  $zipReader->close();

  # Create XPath context
  my $xpath = XML::LibXML::XPathContext->new($doc);
  my $xmlns = $Spreadsheet::Nifty::XLSX::namespaces->{main};
  $xpath->registerNs('excel', $xmlns);

  # Update <sheetProtection/>, if it exists
  my ($sheetProtection) = $xpath->findnodes('/excel:worksheet/excel:sheetProtection');
  if (defined($sheetProtection))
  {
    # Remove any password
    for my $name (qw(password algorithmName hashValue saltValue spinCount))
    {
      $sheetProtection->removeAttribute($name);
    }

    # Unlock the sheet
    $sheetProtection->setAttribute('sheet', '0');
  }

  # Re-serialize worksheet member
  my $member = $fileReader->{zipPackage}->_getMember($rel->{partname});
  $member->contents($doc->toString());

  return 1;
}

1;
