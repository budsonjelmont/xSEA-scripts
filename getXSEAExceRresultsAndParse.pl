#!/usr/bin/perl -w

use diagnostics;
use strict;
use FileHandle;
use warnings;
use Win32::OLE qw(in with);
use Win32::OLE::Const 'Microsoft Office .* Object Library';
use Spreadsheet::ParseExcel;

my $resultsPath = $ARGV[0];
my $timepoint1 = $ARGV[1];
my $timepoint2 = $ARGV[2];

opendir(DIR, "$resultsPath") or die $!;
my ($file1) = grep(/gsea_report_for_.*timepoint${timepoint1}.*\.xls/i, readdir(DIR));
opendir(DIR, "$resultsPath") or die $!;
my ($file2) = grep(/gsea_report_for_.*timepoint${timepoint2}.*\.xls/i, readdir(DIR));

my $file1Output = "xSEAresults_timepoint$timepoint1.txt";
my $file2Output = "xSEAresults_timepoint$timepoint2.txt";

my %files = ($file1 => $file1Output, $file2 => $file2Output);

#Parse excel file and return a tab-delimited file that can be brought back into FM
foreach my $key (keys %files){
	my $output = $files{$key};
	my $file = $key;
	if(open(OUT, "> $output")) {
		my $parser = Spreadsheet::ParseExcel->new();
		my $workbook = $parser->parse($resultsPath . "\\" . $file);
		if ( !defined $workbook ) {
			die $parser->error(), ".\n";
		}
		foreach my $worksheet (@{$workbook->{Worksheet}}) { # looping through worksheets
			for(my $row = $worksheet->{MinRow}; defined $worksheet->{MaxRow} && $row <= $worksheet->{MaxRow}; $row++) { # loop through rows
				my @line;
				for(my $column = $worksheet->{MinCol}; defined $worksheet->{MaxCol} && $column <= $worksheet->{MaxCol}; $column++) { # loop through columns
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
		print "Error: Could not write to output file '$output'\n";
	}
}