#!/usr/bin/perl -w

use diagnostics;
use strict;
use FileHandle;
use warnings;

use Win32::OLE;
my $wsh = new Win32::OLE 'WScript.Shell';
my $desktop_path = $wsh->SpecialFolders('Desktop');

my $foldernamebase = $ARGV[0];
my $condition1 = $ARGV[1];
my $condition2 = $ARGV[2];

opendir(DESKTOP, "$desktop_path") or die $!;
my ($foldername) = grep(/$foldernamebase/i, readdir(DESKTOP));
my $folderpath = "$desktop_path/$foldername";

opendir(DIR, "$folderpath") or die $!;
my ($file1) = grep(/gsea_report_for_.*${condition1}.*\.xls/i, readdir(DIR));
opendir(DIR, "$folderpath") or die $!;
my ($file2) = grep(/gsea_report_for_.*${condition2}.*\.xls/i, readdir(DIR));

my $file1Output = "$desktop_path/xSEAresults_$condition1.txt";
my $file2Output = "$desktop_path/xSEAresults_$condition2.txt";

my %files = ($file1 => $file1Output, $file2 => $file2Output);

#Parse excel file and return a tab-delimited file that can be brought back into FM
#Note: the "excel" file is actually just a tab-delimited file with the xls extension, so all this script really does
#is rename the file and make a copy on the desktop
foreach my $key (keys %files){
	my $output = $files{$key};
	my $file = $folderpath."/".$key;
	if(open(OUT, "> $output")) {
		open IN, $file or die "Couldn't open file: $!";
		while( my $line = <IN>)  {
			if($. > 1){	#don't print first row because we don't need to pull headers into FM
				print OUT $line;
			}
		}
		close(IN);
		close(OUT);
	}
	else {
		print "Error: Could not write to output file '$output'\n";
	}
}