#!/usr/bin/perl -w
####This script opens and modifies an exported expression set from Filemaker and then resaves it as .gct file for use in GSEA

use diagnostics;
use strict;
use Win32::OLE qw(in with);
use Win32::OLE::Const 'Microsoft Excel';

my $filepath = $ARGV[0];
my $replicates = $ARGV[1];

my $outfilepath = substr($filepath,0,length($filepath)-3)."gct";

#stop if errors
$Win32::OLE::Warn = 3;   

# get already active Excel application or open new
my $Excel = Win32::OLE->GetActiveObject('Excel.Application')	
    || Win32::OLE->new('Excel.Application', 'Quit');  
	
# open Excel file
my $Book = $Excel->Workbooks->Open($filepath); 

# select worksheet
my $Sheet = $Book->Worksheets(1);

#get total # of rows used
my $Tot_Rows=$Sheet->UsedRange->Rows->{'Count'}; 
print $Tot_Rows."\n";
#delete unused columns
for(my $i = $replicates+3 ; $i<=18 ; $i++){
	print $i."\n";
	$Sheet->Cells(1, $replicates + 3)->EntireColumn->Delete;
}
for(my $i = ($replicates * 2)+3 ; $i < $replicates+19 ; $i++){
	print $i."\n";
	$Sheet->Cells(1, ($replicates * 2)+3 )->EntireColumn->Delete;
}
#insert two rows for header data
$Sheet->Cells(1,1)->EntireRow->Insert;
$Sheet->Cells(1,1)->EntireRow->Insert;

$Sheet->Cells(1,1)->{'Value'} = "#1.2";
$Sheet->Cells(2,1)->{'Value'} = $Tot_Rows-1;
$Sheet->Cells(3,1)->{'Value'} = "NAME";
$Sheet->Cells(3,2)->{'Value'} = "Description";

#save the file as a .gct
print $outfilepath."\n";
$Excel->{DisplayAlerts} = 0;
#$Book->SaveAs($outfilepath, xlCurrentPlatformText);
$Book->SaveAs($outfilepath, xlTextFormat);
$Book -> close();
$Excel->{DisplayAlerts} = 1;