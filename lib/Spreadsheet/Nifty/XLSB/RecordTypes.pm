#!/usr/bin/perl -w
use warnings;
use strict;

package Spreadsheet::Nifty::XLSB::RecordTypes;

use constant
{
  SST_ITEM             => 0x0013,  # BrtSSTItem
  FONT                 => 0x002B,  # BrtFont
  FMT                  => 0x002C,  # BrtFmt
  FILL                 => 0x002D,  # BrtFill
  BORDER               => 0x002E,  # BrtBorder
  XF                   => 0x002F,  # BrtXF
  STYLE                => 0x0030,  # BrtStyle
  BEGIN_COLOR_PALETTE  => 0x01D9,  # BrtBeginColorPalette
  END_COLOR_PALETTE    => 0x01DA,  # BrtEndColorPalette
  INDEXED_COLOR        => 0x01DB,  # BrtIndexedColor
  BEGIN_INDEXED_COLORS => 0x0235,  # BrtBeginIndexedColors
  END_INDEXED_COLORS   => 0x0236,  # BrtEndIndexedColors
  BEGIN_MRU_COLORS     => 0x0239,  # BrtBeginMRUColors
  END_MRU_COLORS       => 0x023A,  # BrtEndMRUColors
  BEGIN_FILLS          => 0x025B,  # BrtBeginFills
  END_FILLS            => 0x025C,  # BrtEndFills
  BEGIN_FONTS          => 0x0263,  # BrtBeginFonts
  END_FONTS            => 0x0264,  # BrtEndFonts
  BEGIN_BORDERS        => 0x0265,  # BrtBeginBorders
  END_BORDERS          => 0x0266,  # BrtEndBorders
  BEGIN_FMTS           => 0x0267,  # BrtBeginFmts
  END_FMTS             => 0x0268,  # BrtEndFmts
  BEGIN_CELL_XFS       => 0x0269,  # BrtBeginCellXFs
  END_CELL_XFS         => 0x026A,  # BrtEndCellXFs
  BEGIN_STYLES         => 0x026B,  # BrtBeginStyles
  END_STYLES           => 0x026C,  # BrtEndStyles
  BEGIN_CELL_STYLE_XFS => 0x0272,  # BrtBeginCellStyleXFs
  END_CELL_STYLE_XFS   => 0x0273,  # BrtEndCellStyleXFs
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
    '43' => 'BrtFont',
    '44' => 'BrtFmt',
    '45' => 'BrtFill',
    '46' => 'BrtBorder',
    '47' => 'BrtXF',
    '48' => 'BrtStyle',
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
   '278' => 'BrtBeginStyleSheet',
   '279' => 'BrtEndStyleSheet',
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
   '473' => 'BrtBeginColorPalette',
   '474' => 'BrtEndColorPalette',
   '475' => 'BrtIndexedColor',
   '476' => 'BrtMargins',
   '477' => 'BrtPrintOptions',
   '478' => 'BrtPageSetup',
   '485' => 'BrtWsFmtInfo',
   '494' => 'BrtHLink',
   '505' => 'BrtBeginDXFs',
   '506' => 'BrtEndDXFs',
   '507' => 'BrtDXF',
   '508' => 'BrtBeginTableStyles',
   '509' => 'BrtEndTableStyles',
   '510' => 'BrtBeginTableStyle',
   '511' => 'BrtEndTableStyle',
   '534' => 'BrtBookProtection',
   '535' => 'BrtSheetProtection',
   '537' => 'BrtPhoneticInfo',
   '550' => 'BrtDrawing',
   '565' => 'BrtBeginIndexedColors',
   '566' => 'BrtEndIndexedColors',
   '569' => 'BrtBeginMRUColors',
   '570' => 'BrtEndMRUColors',
   '572' => 'BrtMRUColor',
   '551' => 'BrtLegacyDrawing',
   '573' => 'BrtBeginDVals',
   '574' => 'BrtEndDVals',
   '603' => 'BrtBeginFills',
   '604' => 'BrtEndFills',
   '611' => 'BrtBeginFonts',
   '612' => 'BrtEndFonts',
   '613' => 'BrtBeginBorders',
   '614' => 'BrtEndBorders',
   '615' => 'BrtBeginFmts',
   '616' => 'BrtEndFmts',
   '617' => 'BrtBeginCellXFs',
   '618' => 'BrtEndCellXFs',
   '619' => 'BrtBeginStyles',
   '620' => 'BrtEndStyles',
   '625' => 'BrtBigName',
   '626' => 'BrtBeginCellStyleXFs',
   '627' => 'BrtEndCellStyleXFs',
   '648' => 'BrtBeginCellIgnoreECs',
   '649' => 'BrtCellIgnoreEC',
   '650' => 'BrtEndCellIgnoreECs',
   '660' => 'BrtBeginListParts',
   '661' => 'BrtListPart',
   '662' => 'BrtEndListParts',
   '678' => 'BrtSheetProtectionIso',
  '1024' => 'BrtRwDescent',
  '1025' => 'BrtKnownFonts',
  '1045' => 'BrtWsFmtInfoEx14',
  '1131' => 'BrtBeginStyleSheetExt14',
  '1132' => 'BrtEndStyleSheetExt14',
  '1142' => 'BrtBeginSlicerStyles',
  '1143' => 'BrtEndSlicerStyles',
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
