#!/usr/bin/perl -w
use warnings;
use strict;

package Spreadsheet::Nifty::XLSB::RecordTypes;

use constant
{
  SST_ITEM => 0x0013,
};

my $xlsbRecordNames =
{
     '0' => 'BrtRowHdr',
     '1' => 'BrtCellBlank',
     '2' => 'BrtCellRk',
     '3' => 'BrtCellError',
     '4' => 'BrtCellBool',
     '5' => 'BrtCellReal',
     '6' => 'BrtCellSt',
     '7' => 'BrtCellIsst',
     '8' => 'BrtFmlaString',
     '9' => 'BrtFmlaNum',
    '10' => 'BrtFmlaBool',
    '11' => 'BrtFmlaError',
    '19' => 'BrtSSTItem',
    '35' => 'BrtFRTBegin',
    '36' => 'BrtFRTEnd',
    '37' => 'BrtACBegin',
    '38' => 'BrtACEnd',
    '39' => 'BrtName',
    '40' => 'BrtIndexRowBlock',
    '42' => 'BrtIndexBlock',
    '60' => 'BrtColInfo',
    '62' => 'BrtCellRString',
    '64' => 'BrtDVal',
   '128' => 'BrtFileVersion',
   '129' => 'BrtBeginSheet',
   '130' => 'BrtEndSheet',
   '131' => 'BrtBeginBook',
   '132' => 'BrtEndBook',
   '133' => 'BrtBeginWsViews',
   '134' => 'BrtEndWsViews',
   '135' => 'BrtBeginBookViews',
   '136' => 'BrtEndBookViews',
   '137' => 'BrtBeginWsView',
   '138' => 'BrtEndWsView',
   '143' => 'BrtBeginBundleShs',
   '144' => 'BrtEndBundleShs',
   '145' => 'BrtBeginSheetData',
   '146' => 'BrtEndSheetData',
   '147' => 'BrtWsProp',
   '148' => 'BrtWsDim',
   '151' => 'BrtPane',
   '152' => 'BrtSel',
   '153' => 'BrtWbProp',
   '155' => 'BrtFileRecover',
   '156' => 'BrtBundleSh',
   '157' => 'BrtCalcProp',
   '158' => 'BrtBookView',
   '159' => 'BrtBeginSst',
   '160' => 'BrtEndSst',
   '176' => 'BrtMergeCell',
   '177' => 'BrtBeginMergeCells',
   '178' => 'BrtEndMergeCells',
   '277' => 'BrtIndexPartEnd',
   '353' => 'BrtBeginExternals',
   '354' => 'BrtEndExternals',
   '355' => 'BrtSupBookSrc',
   '357' => 'BrtSupSelf',
   '362' => 'BrtExternSheet',
   '390' => 'BrtBeginColInfos',
   '391' => 'BrtEndColInfos',
   '427' => 'BrtShrFmla',
   '461' => 'BrtBeginConditionalFormatting',
   '462' => 'BrtEndConditionalFormatting',
   '463' => 'BrtBeginCFRule',
   '464' => 'BrtEndCFRule',
   '476' => 'BrtMargins',
   '477' => 'BrtPrintOptions',
   '478' => 'BrtPageSetup',
   '485' => 'BrtWsFmtInfo',
   '494' => 'BrtHLink',
   '534' => 'BrtBookProtection',
   '535' => 'BrtSheetProtection',
   '537' => 'BrtPhoneticInfo',
   '550' => 'BrtDrawing',
   '551' => 'BrtLegacyDrawing',
   '573' => 'BrtBeginDVals',
   '574' => 'BrtEndDVals',
   '625' => 'BrtBigName',
   '648' => 'BrtBeginCellIgnoreECs',
   '649' => 'BrtCellIgnoreEC',
   '650' => 'BrtEndCellIgnoreECs',
   '660' => 'BrtBeginListParts',
   '661' => 'BrtListPart',
   '662' => 'BrtEndListParts',
   '678' => 'BrtSheetProtectionIso',
  '1024' => 'BrtRwDescent',
  '1045' => 'BrtWsFmtInfoEx14',
  '2071' => 'BrtAbsPath15',
  '2091' => 'BrtWorkBookPr15',
  '3072' => 'BrtUid',
  '3073' => 'brtRevisionPtr',
  '5095' => 'BrtBeginCalcFeatures',
  '5096' => 'BrtEndCalcFeatures',
  '5097' => 'BrtCalcFeature',
};

sub name($)
{
  my $class = shift();
  my ($num) = @_;

  return $xlsbRecordNames->{sprintf("%d", $num)};
}

1;
