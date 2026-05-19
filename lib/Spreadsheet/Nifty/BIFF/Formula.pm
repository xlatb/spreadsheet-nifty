#!/usr/bin/perl -w
use warnings;
use strict;

package Spreadsheet::Nifty::BIFF::Formula;

use Spreadsheet::Nifty::BIFF::FormulaTokens;

# === Class methods ===

sub new($$)
{
  my $class = shift();
  my ($maxColCount, $maxRowCount) = @_;

  my $self = {};
  $self->{tokens} = [];
  $self->{maxColCount} = $maxColCount;
  $self->{maxRowCount} = $maxRowCount;

  bless($self, $class);
  return $self;
}

# === Instance methods ===

sub appendToken($)
{
  my $self = shift();
  my ($t) = @_;

  push(@{$self->{tokens}}, $t);

  return;
}

sub unparse($)
{
  my $self = shift();
  my ($context) = @_;

  my $tokens = $self->{tokens};

  my $stack = [];

  for my $t (@{$tokens})
  {
    #printf("Formula->unparse(): "); Spreadsheet::Nifty::BIFF::FormulaTokens::dumpToken($t);

    if ($t->{class} == Spreadsheet::Nifty::BIFF::FormulaTokens::TOKEN_CLASS_BASIC)  # Basic tokens
    {
      if ($t->{id} == 0x01)
      {
        # tExp - References a shared formula on the current sheet
        (!defined($context->{getSharedFormula})) && die("Found reference to shared formula but context has no shared formula getter");
        my $sharedFormula = $context->{getSharedFormula}->($t->{exp}->{col}, $t->{exp}->{row});
        (!defined($sharedFormula)) && return undef;  # No such shared formula?

        my $colOffset = $context->{col} - $t->{exp}->{col};
        my $rowOffset = $context->{row} - $t->{exp}->{row};

        # Create a new context for the shared formula
        my $sharedContext = { %{$context} };
        $sharedContext->{col} = ($t->{exp}->{col} + $colOffset) % $self->{maxColCount};
        $sharedContext->{row} = ($t->{exp}->{row} + $rowOffset) % $self->{maxRowCount};

        # Recurse into shared formula
        return $sharedFormula->unparse($sharedContext);
      }
      elsif (($t->{id} >= 0x03) && ($t->{id} <= 0x11))
      {
        # Binary operators
        CORE::state $ops = ['+', '-', '*', '/', '^', '&', '<', '<=', '=', '>=', '>', '<>', ' ', ',', ':'];
        my $args = [ splice(@{$stack}, -2) ];
        my $str = $args->[0] . $ops->[$t->{id} - 0x03] . $args->[1];
        push(@{$stack}, $str);
      }
      elsif (($t->{id} >= 0x12) && ($t->{id} <= 0x13))
      {
        # Prefix unary operators
        CORE::state $ops = ['+', '-'];
        my $arg = pop(@{$stack});
        my $str = $ops->[$t->{id} - 0x12] . $arg;
        push(@{$stack}, $str);
      }
      elsif ($t->{id} == 0x14)
      {
        # tPercent - Postfix unary operator
        my $arg = pop(@{$stack});
        push(@{$stack}, "${arg}%");
      }
      elsif ($t->{id} == 0x15)
      {
        # tParen - Enclosing parentheses
        my $arg = pop(@{$stack});
        push(@{$stack}, "(${arg})");
      }
      elsif ($t->{id} == 0x16)
      {
        # tMissArg - Missing argument
        push(@{$stack}, '');
      }
      elsif ($t->{id} == 0x17)
      {
        # tStr - String literal
        my $str = "\"" . ($t->{str} =~ s/"/""/r) . "\"";
        push(@{$stack}, $str);
      }
      elsif ($t->{id} == 0x19)
      {
        # tAttr - Special attributes
        # TODO: tAttrSpace
        if ($t->{attr}->{flags} == 0x10)
        {
          # tAttrSum - Special-case encoding for SUM() with one argument
          my $arg = pop(@{$stack});
          push(@{$stack}, "SUM(${arg})");
        }
      }
      elsif ($t->{id} == 0x1D)
      {
        # tBool - Boolean literal
        push(@{$stack}, $t->{value} ? 'TRUE' : 'FALSE');
      }
      elsif (($t->{id} == 0x1E) || ($t->{id} == 0x1F))
      {
        # tInt - 16-bit unsigned integer literal
        # tNum - 64-bit IEEE float literal
        push(@{$stack}, "" . $t->{value});
      }
      else
      {
        die(sprintf("Unhandled basic token 0x%02X", $t->{id}));
      }
    }
    else  # Classified tokens
    {
      if ($t->{id} == 0x01)
      {
        # tFunc - Function call with fixed number of args
        my $function = Spreadsheet::Nifty::BIFF::Functions::getFunctionById($t->{func}->{func});
        (!defined($function)) && die(sprintf("Unknown function 0x%02X", $t->{func}->{func}));

        (($function->{minArgs} < 0) || ($function->{minArgs} != $function->{maxArgs})) && die(sprintf("Function '%s' encoded without arg count", $function->{name}));
        my $argCount = $function->{minArgs};

        my $args = [ splice(@{$stack}, -$argCount) ];

        my $str = $function->{name} . '(' . join(',', @{$args}) . ')';
        push(@{$stack}, $str);
      }
      elsif ($t->{id} == 0x02)
      {
        # tFuncVar - Function call with variable number of args
        if ($t->{func}->{func} == 0xFF)
        {
          # Special case for user-defined or future function? First arg is tName with name of the function.
          my $args = [ splice(@{$stack}, -$t->{func}->{argCount}) ];
          my $name = shift(@{$args});
          $name =~ s#^_xlfn\.##;  # Remove future function prefix
          my $str = $name . '(' . join(',', @{$args}) . ')';
          push(@{$stack}, $str);
        }
        else
        {
          # Built-in function with variable number of args
          my $function = Spreadsheet::Nifty::BIFF::Functions::getFunctionById($t->{func}->{func});
          (!defined($function)) && die(sprintf("Unknown function 0x%02X", $t->{func}->{func}));
          # TODO: Check arg count vs function maxArgs

          my $args = ($t->{func}->{argCount}) ? [ splice(@{$stack}, -$t->{func}->{argCount}) ] : [];
          my $str = $function->{name} . '(' . join(',', @{$args}) . ')';
          push(@{$stack}, $str);
        }
      }
      elsif ($t->{id} == 0x03)
      {
        # tName
        #printf("tName index %d\n", $t->{index});
        (!defined($context->{sheet})) && die("No sheet context while unparsing tName token");
        my $label = $context->{sheet}->{workbook}->getLabel($t->{index} - 1);
        push(@{$stack}, $label->{name});
      }
      elsif ($t->{id} == 0x04)
      {
        # tRef - 2d cell reference stored as absolute
        my $str = ((!$t->{ref}->{relCol}) ? '$' : '') . Spreadsheet::Nifty::Utils->colIndexToString($t->{ref}->{col}) . ((!$t->{ref}->{relRow}) ? '$' : '') . int($t->{ref}->{row} + 1);
        push(@{$stack}, $str);
      }
      elsif ($t->{id} == 0x05)
      {
        # tArea - 2d area reference stored as absolute
        my $minCol = $t->{ref}->{relMinCol} ? ($t->{ref}->{minCol} % $self->{maxColCount}) : $t->{ref}->{minCol};
        my $maxCol = $t->{ref}->{relMaxCol} ? ($t->{ref}->{maxCol} % $self->{maxColCount}) : $t->{ref}->{maxCol};
        my $minRow = $t->{ref}->{relMinRow} ? ($t->{ref}->{minRow} % $self->{maxRowCount}) : $t->{ref}->{minRow};
        my $maxRow = $t->{ref}->{relMaxRow} ? ($t->{ref}->{maxRow} % $self->{maxRowCount}) : $t->{ref}->{maxRow};
        #print main::Dumper($context->{col}, $context->{row}, $t, $minCol, $minRow, $maxCol, $maxRow);
        my $str = sprintf("%s%s%s%s:%s%s%s%s",
         $t->{ref}->{relMinCol} ? '' : '$',
         Spreadsheet::Nifty::Utils->colIndexToString($minCol),
         $t->{ref}->{relMinRow} ? '' : '$',
         int($minRow + 1),
         $t->{ref}->{relMaxCol} ? '' : '$',
         Spreadsheet::Nifty::Utils->colIndexToString($maxCol),
         $t->{ref}->{relMaxRow} ? '' : '$',
         int($maxRow + 1));
        push(@{$stack}, $str);
      }
      elsif ($t->{id} == 0x0C)
      {
        # tRefN - 2D cell reference stored as a relative offset
        my $col = $t->{ref}->{relCol} ? ($context->{col} + $t->{ref}->{col}) % $self->{maxColCount} : $t->{ref}->{col};
        my $row = $t->{ref}->{relRow} ? ($context->{row} + $t->{ref}->{row}) % $self->{maxRowCount} : $t->{ref}->{row};
        #print main::Dumper('tRefN', $context->{col}, $context->{row}, $t, $col, $row);

        my $str = ((!$t->{ref}->{relCol}) ? '$' : '') . Spreadsheet::Nifty::Utils->colIndexToString($col) . ((!$t->{ref}->{relRow}) ? '$' : '') . int($row + 1);
        push(@{$stack}, $str);
      }
      elsif ($t->{id} == 0x0D)
      {
        # tAreaN - 2D area reference stored as a relative offset
        my $minCol = $t->{ref}->{relMinCol} ? ($context->{col} + $t->{ref}->{minCol}) % $self->{maxColCount} : $t->{ref}->{minCol};
        my $maxCol = $t->{ref}->{relMaxCol} ? ($context->{col} + $t->{ref}->{maxCol}) % $self->{maxColCount} : $t->{ref}->{maxCol};
        my $minRow = $t->{ref}->{relMinRow} ? ($context->{row} + $t->{ref}->{minRow}) % $self->{maxRowCount} : $t->{ref}->{minRow};
        my $maxRow = $t->{ref}->{relMaxRow} ? ($context->{row} + $t->{ref}->{maxRow}) % $self->{maxRowCount} : $t->{ref}->{maxRow};
        #print main::Dumper($context->{col}, $context->{row}, $t, $minCol, $minRow, $maxCol, $maxRow);
        my $str = sprintf("%s%s%s%s:%s%s%s%s",
         $t->{ref}->{relMinCol} ? '' : '$',
         Spreadsheet::Nifty::Utils->colIndexToString($minCol),
         $t->{ref}->{relMinRow} ? '' : '$',
         int($minRow + 1),
         $t->{ref}->{relMaxCol} ? '' : '$',
         Spreadsheet::Nifty::Utils->colIndexToString($maxCol),
         $t->{ref}->{relMaxRow} ? '' : '$',
         int($maxRow + 1));
        push(@{$stack}, $str);
      }
      elsif ($t->{id} == 0x1A)
      {
        # tRef3d
        ...;  # TODO: Needs to look up external sheet name(s)
      }
      else
      {
        die(sprintf("Unhandled classified token: id 0x%02X", $t->{id}));
      }
    }

    #print main::Dumper('STACK', $stack);
  }

  my $stackLength = scalar(@{$stack});
  ($stackLength != 1) && die("unparseTokens(): Expected stack length 1, got $stackLength");

  return $stack->[0];
}

1;
