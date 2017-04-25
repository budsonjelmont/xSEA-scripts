#!/usr/local/bin/perl

use strict;
use Spreadsheet::ParseExcel;

my $inputFile    = shift;
my $replicates	= shift;
my $outputFile  = substr($inputFile,0,length($inputFile)-3)."gct";

if(!-e $inputFile || !$outputFile) {
    print qq~
Usage: excelToTabDelim inputFile outputFile
    Parses an excel input file into a tab delimited out file
~;
    exit 1;
}

if(open(OUT, "> $outputFile")) {
    my $workbook = Spreadsheet::ParseExcel::Workbook->Parse($inputFile);
    foreach my $worksheet (@{$workbook->{Worksheet}}) { # looping through worksheets
		#construct header
		my $Tot_Rows= $worksheet->{MaxRow};
		print "Total rows: ".$Tot_Rows."\n";
		print OUT "#1.2\n";
		print OUT ($Tot_Rows)."\t".($replicates*2)."\n";
		print OUT "NAME\tDescription\t";
		my @line;
        for(my $column = 2; $column <= $replicates+1 ; $column++) { # loop through columns
			my $value = $worksheet->{Cells}[0][$column] ? $worksheet->{Cells}[0][$column]->Value : "";
			$value =~ s/^\s+//;
			$value =~ s/\s+$//;
			$value =~ s/[\r\n]+//g;
			$value = " " if $value ne "0" && !$value;
			push(@line, $value);
		}
        for(my $column = 18; $column <= 17+$replicates ; $column++) { # loop through columns
			my $value = $worksheet->{Cells}[0][$column] ? $worksheet->{Cells}[0][$column]->Value : "";
			$value =~ s/^\s+//;
			$value =~ s/\s+$//;
			$value =~ s/[\r\n]+//g;
			$value = " " if $value ne "0" && !$value;
			push(@line, $value);
		}
		$" = "\t";
		print OUT "@line\n";
		#loop through the rest of the rows
        for(my $row = 1; defined $worksheet->{MaxRow} && $row <= $worksheet->{MaxRow}; $row++) { # loop through rows
            my @line;
            for(my $column = 0; $column <= $replicates+1 ; $column++) { # loop through columns
				my $value = $worksheet->{Cells}[$row][$column] ? $worksheet->{Cells}[$row][$column]->Value : "";
				$value =~ s/^\s+//;
				$value =~ s/\s+$//;
				$value =~ s/[\r\n]+//g;
				$value = " " if $value ne "0" && !$value;
				push(@line, $value);
			}
			for(my $column = 18; $column <= 17+$replicates ; $column++) { # loop through columns
				my $value = $worksheet->{Cells}[$row][$column] ? $worksheet->{Cells}[$row][$column]->Value : "";
				$value =~ s/^\s+//;
				$value =~ s/\s+$//;
				$value =~ s/[\r\n]+//g;
				$value = " " if $value ne "0" && !$value;
				push(@line, $value);
			}
            $" = "\t";
            print OUT "@line\n";
        }
    }
    close(OUT);
}
else {
    print "Error: Could not write to output file '$outputFile'\n";
}