#!/usr/bin/perl

#use strict;
#use warnings;

use Cwd;
use Getopt::Long;
use File::Basename;
use Data::Dumper;

our $php_location = "C:\\PHP\\php.exe";
our $lead_sheet_name = "Leads";

our $subsystem_report_type = "Subsystem";

our $staging_worksheet_name = "Staging Worksheet";
our $scratch_staging_worksheet_name = "Scratch Staging Worksheet";

our $lead_column;
our $assignee_column;
our $send_once_column;
our $format_columm;
# our $parameter_col;

our $html_file_name = "toemail.html";
our $sending_file_name = "embed.html";
our $techlead_file_name = "techlead.txt";
our $saveashtml_macro_name = "Save_Staging";
our $blatlocation = "C:\\Blat\\full\\blat.exe";

our $parameters_sheet_name = "Parameters";
our %assignee_to_lead; # keys are the assignees in $assignee_column in $lead_sheet_name


# list of scratch worksheets to be deleted at beginning of script, and at end
# our @TempDeleteSheetNames = ("Temp", "Temp2", "Copy of Staging Worksheet", "Copy of Staging", "Staging Worksheet", "Scratch Staging Worksheet");
our @TempDeleteSheetNames = ("Temp", "Temp2", "Sheet1", "Copy of Staging Worksheet", "Copy of Staging", "Scratch Staging Worksheet");

our $ok_to_process_all = 1;
our $preview_members = 'c_jhenk, mhuh';
our %sendtohash;
# send $sending_file_name in an email to recipients
our @recipients; # names of who to send to


use Win32::OLE qw(in with);
use Win32::OLE::Const 'Microsoft Excel';

my $PROGRAM = fileparse($0);

# get already active Excel application or open new
our $Excel = Win32::OLE->GetActiveObject('Excel.Application')
    || Win32::OLE->new('Excel.Application', 'Quit');

#don't display alerts
$Excel->{DisplayAlerts} = False;

## Parse command line arguments
Getopt::Long::Configure('pass_through');
GetOptions(
    "help"              => \my $help_opt,
    "file=s"            => \my $file_opt,
    "config=s"          => \my $config_opt,
    "verbose"           => \my $verbose,
    "xdebug"            => \my $debug,
);

if (!$file_opt) {
   $file_opt = cwd()."\\CR_Notification_Configuration.xlsm";
   print "*** -file option not set - setting to default - $file_opt\n";
} else {
	$file_opt = $file_opt = cwd()."\\".$file_opt;
}

if (!$config_opt) {
   $config_opt = "Config Instances";
   print "*** -config option not set - setting to default - $config_opt\n";
}

if ($help_opt) {
   usage();
}

my %techlead_hash;
load_techlead_hash();			# hash to be loaded with full names that are ambiguous in ph (with the correct email addresses) from local file

print "$file_opt\n";
our $Book = $Excel->Workbooks->Open($file_opt);
makeLeadHash();

our	$Sheet = $Book->Worksheets($config_opt);
# open config worksheet
# find the last row and column in the sheet
our $LastRow = $Sheet->UsedRange->Find({What=>"*",
	SearchDirection=>xlPrevious,
	SearchOrder=>xlByRows})->{Row};
our $LastCol = $Sheet->UsedRange->SpecialCells(xlCellTypeLastCell)->{Column};

for ($i = 1; $i <= $LastCol_Config; $i++) {
	if ($Sheet->Cells(1, $i)->{'Value'} eq 'Send Once') {
		$send_once_column = $i;
	} elsif ($Sheet->Cells(1, $i)->{'Value'} eq 'Format') {
		$format_column = $i;
#	} elsif ($Sheet->Cells(1, $i)->{'Value'} eq 'Parameter') {
#		$parameter_column = $i;
	}
}

empty_dir();

# temp paramater storage
our %temphash=();
our $tempvalue;

# contains the names of the parameters(these are the keys in the hash)
our @param_names;

# array of parameters, this is an array of hashes
our @parameters;
our $col;
our $row;

# store the names of the parameters from the first line
for (my $paramcell=1; $paramcell <= $LastCol; $paramcell++) {
	push (@param_names, $Sheet->Cells(1, $paramcell)->{'Value'}) ;
}

# search through the sheet, one row at
# a time and add parameters to %temphash
# then push() these onto the @config_paramater

for ($row= 2; $row <= $LastRow; $row++) {
	#clear temp params

	#skip undefined rows, undefined rows have no configuration name
	next unless defined $Sheet->Cells($row,1)->{'Value'};
	for ($col=1; $col <= $LastCol; $col++) {
		# add all defined cells to %temphash
		$tempvalue = $Sheet->Cells($row, $col)->{'Value'};
		if (defined $tempvalue) {
			$temphash{$param_names[$col-1] } = $tempvalue;
		}
	}
	print "\n";
	# add one row of parameters to the @parameters array
	push @parameters, {%temphash};
	print "Row #$row is stored as configuration.\n";
}

if (ParseParamsAndEmail(@parameters)) {
	print "\n********Success***********\n";
}

# clean up after ourselves
cleanUpExcel();
$Book->Save();
$Book->Close();
empty_dir();

print "Script done\n";
sleep (3);

## ======================================================================
##
## Routine - load_techlead_hash
## input - none
## output - none
## returns - none
##
## Reads local file with full names and corresponding email addresses -
## contents to go into global hash - full names are intended to be
## ambiguous names in ph - ugly way to avoid fruitless searches.
##
## ======================================================================
sub load_techlead_hash {
	my @tech_ary = ();
	open TECHLEAD, "<$techlead_file_name";
	while (<TECHLEAD>) {
		@tech_ary = split(/,\s/, $_);
		$techlead_hash{$tech_ary[0]} = $tech_ary[1];
	}
	close TECHLEAD;
}

## ======================================================================
##
## Routine - usage
## input - none
## output - Help Screen dump
## returns - none
##
## ======================================================================
sub usage {
   print "\n*** Usage: $PROGRAM -h\n";
   print "           $PROGRAM <-f xlsm-file-name> <-c instances-sheet> <-v> <-x>\n";
   print "\n";
   print "    -h help (print this screen\n";
   print "    -f xlsm file name                  current setting: $file_opt\n";
   print "    -c instances page on spreadsheet   current setting: $config_opt\n";
   print "    -v (verbose)                       current setting: $verbose\n";
   print "    -x (debug)                         current setting: $debug\n";

   exit();
}

sub empty_dir() {
	my $temp = $html_file_name;
	$temp =~ s{\.html}{};
	$temp .= "_files";
	my @files = <$temp/*>;
	foreach $file (@files) {
		unlink($file);
	}
	@files = <*.png>;
	foreach $file (@files) {
		unlink($file);
	}
}

sub EmbedPictureInHtml {
	#looks for png files in cwd\*.png
	#and embeds them into the $sending_file_name file
	#
	my @unsorted_files = <*.png>;
	my @files = sort @unsorted_files;

	$mystring = Dumper(@files);
	print "files is:\n$mystring\n";

	open FILE, "<$sending_file_name";

	my @lines = <FILE>;
	my @newlines;
	foreach(@lines)	{
		if (m{<body>}) {
			print "\nadding image(s) to html\n";
			foreach $filename (@files) {
				if (defined $ARGV[1] and $ARGV[1] =~ m{preview}) {
					$_ .= '<br><img src=$filename alt="" >';
				} else {
					$_ .= "<br><img src =\"cid:$filename\" alt=\"\">\n";
				}
			}
		} elsif (/align=center/) {
			s/align=center/align=left/;
		}
		push @newlines, $_;
	}
	close FILE;

	open FILE, ">$sending_file_name";

	foreach(@newlines) {
		print FILE $_;
	}

	close FILE;

	return 1;
}

sub addContentToHtml {
	#Takes the instance_parameters above_table content and below_table_content and adds it to the html
	my $above_email_content = shift();
	my $below_email_content = shift();

	open FILE, "<$sending_file_name";

	my @lines = <FILE>;
	my @newlines;
	foreach(@lines)	{
		if (m{<body>}) {
			$_ .= "<br>$above_email_content<br>";
		}
		if (m{<\/body>}) {
			$_ .= "<br>$below_email_content<br><br>";
		}
		push	@newlines, $_;
	}

	close FILE;
	open FILE, ">$sending_file_name";

	foreach(@newlines) {
		print FILE $_;
	}

	close FILE;
	return 1;
}

sub QueryPrism {
	#calls some php/gets info from url about a query ID and stores it
	#in an array of hashes, each subarray contains a row that will be put into an excel sheet
	#uses the argument as a query #

	print "Querying for id $_[0]...\n\n Start time -> ";
	printTime();
	system("$php_location", "querygetter.php", "$_[0]" );
	print "\nEnd time ->";
	printTime();
	my @tableinfo;

	#holds overall results for the queryprism subroutine
	my @results;

	#****** for now, read from file
	open FILE, "crlist.txt" or die $!;
	my @lines;
	push @lines, $_ foreach (<FILE>);

	my $line = join //, @lines;

	#*****

	#delimeter constants
	$char1 = chr 1;
	$char2 = chr 2;
	$char3 = chr 3;

	#split the one big line into an array : @tableinfo

	@tableinfo = split /$char2/, $line;

	my @temparray; #temp array to hold subarray

	my %temphash;
	my $tempkey;
	my $tempvalue;

	# put infomation into an array of hashes
	foreach $CR (@tableinfo) {
		my @info = split /$char1/, $CR; # split each CR into an array of data ( TITLE \x3 VALUE)
		foreach $chunk (@info) {
			($tempkey, $tempvalue) = split /$char3/, $chunk; # split each data chunk into a key and value
			if (defined $tempvalue) { # put data into a hash
				$temphash{$tempkey} = $tempvalue;
			} else { # or an empty string if undefined
				$temphash{$tempkey} = "";
			}
		}
		#push temp hash onto the results array
		push @results, {%temphash};
	}

	return @results;
}

sub GetCRDataFields {
	#given a query id, look for one CR in the query and get all of the data fields.
	#return them in an array
	my $num_args = scalar @_;
	if ($num_args != 1) {
		print "\nCR data fields cannot be loaded, incorrect number of arguments: $num_args.\n";
		return 0;
	}
	print "calling php script..\n";
	#call php script that loads cr data fields into datafields.txt
	system( "C:\\php\\php.exe crdatafields.php $_[0]");

	open FILE, "datafields.txt"; #open file
	my $datafields_raw = <FILE>; #read in data from file
	my @data_fields = split /\x3/, $datafields_raw; #split into different data fields
	close FILE;

	return @data_fields;
}

sub makeLeadHash {
	# creates a lead-assignee hash
	# lead1->assignee1,assignee2
	# lead2->assignee3
	#
	my @LeadAry;
    my $row;
    my $assignee;
    my $lead;
    my $result;
	my $assignee_column = 1;
	my $lead_column = 2;

	my $lead_sheet = $Book->Worksheets("$lead_sheet_name");
	my $last_lead_row = $lead_sheet->UsedRange->Find({What=>"*",
	SearchDirection=>xlPrevious,
	SearchOrder=>xlByRows})->{Row};

	for ($row = 1; $row <= $last_lead_row; $row++) {
		$assignee = $lead_sheet->Cells($row, $assignee_column)->{'Value'};
		$assignee =~ s/\s//g;   # remove spaces
		@LeadAry = split /,/, $lead_sheet->Cells($row, $lead_column)->{'Value'};
		foreach $result (@LeadAry) {			#find unique names of PLAssignees
			$lead =~ s/\s//g;   # remove spaces
			$assignee_to_lead{$assignee} = $lead;
		}
	}
}

#sub WriteColumnParameters { # writes to column
#    print "writing to drop down column params\n";
#    my $row = 2;
    # alphabetise
#    my @params = sort @_;
#    foreach $field (@params) {
#        $parameters_sheet->Cells($row, $parameter_col)->{Value} = $field;
#        $row++;
#    }
#    print "\nDone writing to parameter column.\n";
#}
#
sub WriteInfoToStagingWorksheet {
	my ($instance_ref, $queryresults_ref) = @_;

    my %local_instance = %$instance_ref;
    my @local_queryresults = @$queryresults_ref;


	$Excel->{DisplayAlerts} = False;
	cleanUpExcel();
	#This sub writes information (in argument) to a staging worksheet
	#Calls a formatting macro in the excel sheet
	#Saves sheet as html file

	my $remove_dupes    = $local_instance{'Remove Dupes'};
	my $graph_only      = $local_instance{'Graph Only'};
	my $graph_type      = $local_instance{'Graph Type'};
	my $formatting      = $local_instance{'Format'};
	my $PLAssignee_flag = $local_instance{'Send to PL Assignee'};
	my $FAAssignee_flag = $local_instance{'Send to FA Assignee'};
	my $TechLead_flag   = $local_instance{'Send to Tech Lead'};
	my $Lead_flag       = $local_instance{'Send to Lead'};

	#create and open a new worksheet named $staging_worksheet_name
	my $staging_worksheet = $Book->WorkSheets->Add;
	$staging_worksheet->{Name} = $staging_worksheet_name;

	print "\nwriting to sheet...\n";
	# get all the values from paramaters
	# that will correspond to the values to put on the tables
	# i.e. which columns are we going to show

	$formatting =~ s/\s//g; #remove whitespace
	my @columns = split /;/, $formatting;

	# put info into staging_worksheet
	my $col_num = 1;
	my $row_num = 1;
	foreach $name (@columns ){ # write names to top line
		$staging_worksheet->Cells($row_num, $col_num)->{Value} = $name;
		$col_num++;
	}
	$row_num++; #now at row 2, col 1

	my %temphash;

	foreach $hash (@local_queryresults) {
		# get each hash out of the array
		# these contain the values from prism
		my %temphash = %{$hash}; #derefence scalar as a hash

		#write the values of current hash to line
		$col_num = 1;
		foreach $name (@columns ) {
			$staging_worksheet->Cells($row_num, $col_num)->{Value} = $temphash{$name};
			$col_num++;
		}
		$row_num++;
	}

	$Excel->Run("Format_Table");

	if ($remove_dupes eq "Y") {
		$Excel->Run("RemoveDupes");
	}

	# format macro &
	# delete Temp sheet if using statistics( aresult of the macro)
	if ($graph_type eq "Individual") {
		$Excel->Run("Make_Scratch_WS_and_Paste");
		$Excel->Run("Get_Stats", "PLAssignee");				# make second graph in Temp w/s
		$Excel->Run("do_Export_Chart_first");				# export .png file to cwd
		if ($graph_only ne 'Y') {
			$Excel->Run("$saveashtml_macro_name");			# This db is desired as to be published
		}
		add_to_recipient_list_from_scratch_sheet($PLAssignee_flag, $FAAssignee_flag, $TechLead_flag, $Lead_flag);
		$Excel->Run("Delete_Scratch_and_Temp_Worksheet");	# delete scratch data w/s
	} elsif ($graph_type eq "Functional Area") {
		$Excel->Run("Make_Scratch_WS_and_Paste");
		$Excel->Run("Get_Stats", "FunctionalArea");			# make second graph in Temp w/s
		$Excel->Run("do_Export_Chart_first");				# export .png file to cwd
		if ($graph_only ne 'Y') {
			$Excel->Run("$saveashtml_macro_name");			# This db is desired as to be published
		}
		add_to_recipient_list_from_scratch_sheet($PLAssignee_flag, $FAAssignee_flag, $TechLead_flag, $Lead_flag);
		$Excel->Run("Delete_Scratch_and_Temp_Worksheet");	# delete scratch data w/s
	} elsif ($graph_type eq "Area") {
		$Excel->Run("Make_Scratch_WS_and_Paste");
		$Excel->Run("Get_Stats", "Area");				# make second graph in Temp w/s
		$Excel->Run("do_Export_Chart_first");				# export .png file to cwd
		if ($graph_only ne 'Y') {
			$Excel->Run("$saveashtml_macro_name");			# This db is desired as to be published
		}
		add_to_recipient_list_from_scratch_sheet($PLAssignee_flag, $FAAssignee_flag, $TechLead_flag, $Lead_flag);
		$Excel->Run("Delete_Scratch_and_Temp_Worksheet");	# delete scratch data w/s
	} elsif ($graph_type eq "Subsystem") {
		$Excel->Run("Make_Scratch_WS_and_Paste");
		$Excel->Run("Get_Stats", "Subsystem");				# make second graph in Temp w/s
		$Excel->Run("do_Export_Chart_first");				# export .png file to cwd
		if ($graph_only ne 'Y') {
			$Excel->Run("$saveashtml_macro_name");			# This db is desired as to be published
		}
		add_to_recipient_list_from_scratch_sheet($PLAssignee_flag, $FAAssignee_flag, $TechLead_flag, $Lead_flag);
		$Excel->Run("Delete_Scratch_and_Temp_Worksheet");	# delete scratch data w/s
	} elsif ($graph_type eq "Both") {			# make two graphs
		$Excel->Run("Make_Scratch_WS_and_Paste");
		$Excel->Run("Get_Stats", "FunctionalArea");				# make second graph in Temp w/s
		$Excel->Run("do_Export_Chart_first");				# export .png file to cwd
		if ($graph_only ne 'Y') {
			$Excel->Run("$saveashtml_macro_name");			# This db is desired as to be published
		}
		add_to_recipient_list_from_scratch_sheet($PLAssignee_flag, $FAAssignee_flag, $TechLead_flag, $Lead_flag);
		$Excel->Run("Delete_Scratch_and_Temp_Worksheet");	# delete scratch data w/s

		$Excel->Run("Make_Scratch_WS_and_Paste");			# put data into scratch data w/s
		$Excel->Run("Get_Stats", "PLAssignee");				# make second graph in Temp w/s
		$Excel->Run("do_Export_Chart_second");				# export .png file to cwd
		if ($graph_only ne 'Y') {
			$Excel->Run("$saveashtml_macro_name");			# This db is desired as to be published
		}
		add_to_recipient_list_from_scratch_sheet($PLAssignee_flag, $FAAssignee_flag, $TechLead_flag, $Lead_flag);
		$Excel->Run("Delete_Scratch_and_Temp_Worksheet");	# delete scratch data w/s

	} elsif ($graph_type eq "A_S_Modem") {		# make two graphs
		print "Modem.\n";
		$Excel->Run("Make_Scratch_WS_and_Paste");			# put data into scratch data w/s
		$Excel->Run("Get_Stats", "Area");					# make second graph in Temp w/s
		$Excel->Run("do_Export_Chart_first");				# export .png file to cwd
		if ($graph_only ne 'Y') {
			$Excel->Run("$saveashtml_macro_name");			# This db is desired as to be published
		}
		add_to_recipient_list_from_scratch_sheet($PLAssignee_flag, $FAAssignee_flag, $TechLead_flag, $Lead_flag);
		$Excel->Run("Delete_Scratch_and_Temp_Worksheet");	# delete scratch data w/s

		$Excel->Run("Make_Scratch_WS_and_Paste");			# put data into scratch data w/s
		$Excel->Run("DeleteRowsModem");						# EXAMPLE DB MANIPULATION - scratch w/s db - deletes all but Modem
		$Excel->Run("Get_Stats", "Subsystem", "Modem");		# make second graph in Temp w/s
		$Excel->Run("do_Export_Chart_second");				# export .png file to cwd
		add_to_recipient_list_from_scratch_sheet($PLAssignee_flag, $FAAssignee_flag, $TechLead_flag, $Lead_flag);
		$Excel->Run("Delete_Scratch_and_Temp_Worksheet");	# delete scratch data w/s

	} elsif ($graph_type eq "A_S_Core") {		# make two graphs
		print "Core.\n";
		$Excel->Run("Make_Scratch_WS_and_Paste");			# put data into scratch data w/s
		$Excel->Run("Get_Stats", "Area");					# make second graph in Temp w/s
		$Excel->Run("do_Export_Chart_first");				# export .png file to cwd
		add_to_recipient_list_from_scratch_sheet($PLAssignee_flag, $FAAssignee_flag, $TechLead_flag, $Lead_flag);
		$Excel->Run("Delete_Scratch_and_Temp_Worksheet");	# delete scratch data w/s

		$Excel->Run("Make_Scratch_WS_and_Paste");			# put data into scratch data w/s
		$Excel->Run("DeleteRowsCore");						# EXAMPLE DB MANIPULATION - scratch w/s db - deletes all but Core
		$Excel->Run("Get_Stats", "Subsystem", "Core");		# make second graph in Temp w/s
		$Excel->Run("do_Export_Chart_second");				# export .png file to cwd
		if ($graph_only ne 'Y') {
			$Excel->Run("$saveashtml_macro_name");				# This db is desired as to be published
		}
		add_to_recipient_list_from_scratch_sheet($PLAssignee_flag, $FAAssignee_flag, $TechLead_flag, $Lead_flag);
		$Excel->Run("Delete_Scratch_and_Temp_Worksheet");	# delete scratch data w/s

	} elsif ($graph_type eq "A_S_All") {					# make two graphs
		print "All.\n";
		$Excel->Run("Make_Scratch_WS_and_Paste");			# put data into scratch data w/s
		$Excel->Run("Get_Stats", "Area");					# make second graph in Temp w/s
		$Excel->Run("do_Export_Chart_first");				# export .png file to cwd
		if ($graph_only ne 'Y') {
			$Excel->Run("$saveashtml_macro_name");			# This db is desired as to be published
		}
		add_to_recipient_list_from_scratch_sheet($PLAssignee_flag, $FAAssignee_flag, $TechLead_flag, $Lead_flag);
		$Excel->Run("Delete_Scratch_and_Temp_Worksheet");	# delete scratch data w/s

		$Excel->Run("Make_Scratch_WS_and_Paste");			# put data into scratch data w/s
		$Excel->Run("Get_Stats", "Subsystem");				# make second graph in Temp w/s
		$Excel->Run("do_Export_Chart_second");				# export .png file to cwd
		add_to_recipient_list_from_scratch_sheet($PLAssignee_flag, $FAAssignee_flag, $TechLead_flag, $Lead_flag);
		$Excel->Run("Delete_Scratch_and_Temp_Worksheet");	# delete scratch data w/s

	} elsif ($graph_type eq "A_S_All_Three") {				# make two graphs
		print "All Three.\n";
		$Excel->Run("Make_Scratch_WS_and_Paste");			# put data into scratch data w/s
		$Excel->Run("Get_Stats", "Area");					# make second graph in Temp w/s
		$Excel->Run("do_Export_Chart_first");				# export .png file to cwd
		if ($graph_only ne 'Y') {
			$Excel->Run("$saveashtml_macro_name");			# This db is desired as to be published
		}
		add_to_recipient_list_from_scratch_sheet($PLAssignee_flag, $FAAssignee_flag, $TechLead_flag, $Lead_flag);
		$Excel->Run("Delete_Scratch_and_Temp_Worksheet");	# delete scratch data w/s

		$Excel->Run("Make_Scratch_WS_and_Paste");			# put data into scratch data w/s
		$Excel->Run("DeleteRowsCore");						# EXAMPLE DB MANIPULATION - scratch w/s db - deletes all but Core
		$Excel->Run("Get_Stats", "Subsystem", "Core");		# make second graph in Temp w/s
		$Excel->Run("do_Export_Chart_second");				# export .png file to cwd
		add_to_recipient_list_from_scratch_sheet($PLAssignee_flag, $FAAssignee_flag, $TechLead_flag, $Lead_flag);
		$Excel->Run("Delete_Scratch_and_Temp_Worksheet");	# delete scratch data w/s

		$Excel->Run("Make_Scratch_WS_and_Paste");			# put data into scratch data w/s
		$Excel->Run("DeleteRowsModem");						# EXAMPLE DB MANIPULATION - scratch w/s db - deletes all but Modem
		$Excel->Run("Get_Stats", "Subsystem", "Modem");		# make second graph in Temp w/s
		$Excel->Run("do_Export_Chart_third");				# export .png file to cwd
		add_to_recipient_list_from_scratch_sheet($PLAssignee_flag, $FAAssignee_flag, $TechLead_flag, $Lead_flag);
		$Excel->Run("Delete_Scratch_and_Temp_Worksheet");	# delete scratch data w/s

	} elsif ($graph_type eq "A_S_All_Subsystem") {		# make two graphs
		print "A_S_All_Subsystem\n";
		$Excel->Run("Make_Scratch_WS_and_Paste");			# put data into scratch data w/s
		$Excel->Run("DeleteRowsCore");						# EXAMPLE DB MANIPULATION - scratch w/s db - deletes all but Core
		$Excel->Run("Get_Stats", "Subsystem", "Core");					# make second graph in Temp w/s
		$Excel->Run("do_Export_Chart_first");				# export .png file to cwd
		add_to_recipient_list_from_scratch_sheet($PLAssignee_flag, $FAAssignee_flag, $TechLead_flag, $Lead_flag);
		$Excel->Run("Delete_Scratch_and_Temp_Worksheet");	# delete scratch data w/s

		$Excel->Run("Make_Scratch_WS_and_Paste");			# put data into scratch data w/s
		$Excel->Run("DeleteRowsModem");						# EXAMPLE DB MANIPULATION - scratch w/s db - deletes all but Modem
		$Excel->Run("Get_Stats", "Subsystem", "Modem");					# make second graph in Temp w/s
		$Excel->Run("do_Export_Chart_second");				# export .png file to cwd
		add_to_recipient_list_from_scratch_sheet($PLAssignee_flag, $FAAssignee_flag, $TechLead_flag, $Lead_flag);
		$Excel->Run("Delete_Scratch_and_Temp_Worksheet");	# delete scratch data w/s

		$Excel->Run("Make_Scratch_WS_and_Paste");			# put data into scratch data w/s
		$Excel->Run("DeleteRowsModemAndCorePositive");		# EXAMPLE DB MANIPULATION - scratch w/s db - deletes Core and Modem
		$Excel->Run("Get_Stats", "Subsystem", "Others");	# make second graph in Temp w/s
		$Excel->Run("do_Export_Chart_third");				# export .png file to cwd
		if ($graph_only ne 'Y') {
			$Excel->Run("$saveashtml_macro_name");			# This db is desired as to be published
		}
		add_to_recipient_list_from_scratch_sheet($PLAssignee_flag, $FAAssignee_flag, $TechLead_flag, $Lead_flag);
		$Excel->Run("Delete_Scratch_and_Temp_Worksheet");	# delete scratch data w/s

	} elsif ($graph_type eq "A_S_All_Four") {		# make two graphs
		print "All Four.\n";
		$Excel->Run("Make_Scratch_WS_and_Paste");			# put data into scratch data w/s
		$Excel->Run("Get_Stats", "Area");					# make second graph in Temp w/s
		$Excel->Run("do_Export_Chart_first");				# export .png file to cwd
		if ($graph_only ne 'Y') {
			$Excel->Run("$saveashtml_macro_name");			# This db is desired as to be published
		}
		add_to_recipient_list_from_scratch_sheet($PLAssignee_flag, $FAAssignee_flag, $TechLead_flag, $Lead_flag);
		$Excel->Run("Delete_Scratch_and_Temp_Worksheet");	# delete scratch data w/s

		$Excel->Run("Make_Scratch_WS_and_Paste");			# put data into scratch data w/s
		$Excel->Run("DeleteRowsModem");						# EXAMPLE DB MANIPULATION - scratch w/s db - deletes all but Core
		$Excel->Run("Get_Stats", "Subsystem", "Modem");		# make second graph in Temp w/s
		$Excel->Run("do_Export_Chart_second");				# export .png file to cwd
		add_to_recipient_list_from_scratch_sheet($PLAssignee_flag, $FAAssignee_flag, $TechLead_flag, $Lead_flag);
		$Excel->Run("Delete_Scratch_and_Temp_Worksheet");	# delete scratch data w/s

		$Excel->Run("Make_Scratch_WS_and_Paste");			# put data into scratch data w/s
		$Excel->Run("DeleteRowsCore");						# EXAMPLE DB MANIPULATION - scratch w/s db - deletes all but Modem
		$Excel->Run("Get_Stats", "Subsystem", "Core");		# make second graph in Temp w/s
		$Excel->Run("do_Export_Chart_third");				# export .png file to cwd
		add_to_recipient_list_from_scratch_sheet($PLAssignee_flag, $FAAssignee_flag, $TechLead_flag, $Lead_flag);
		$Excel->Run("Delete_Scratch_and_Temp_Worksheet");	# delete scratch data w/s

		$Excel->Run("Make_Scratch_WS_and_Paste");			# put data into scratch data w/s
		$Excel->Run("DeleteRowsModemAndCorePositive");		# EXAMPLE DB MANIPULATION - scratch w/s db - deletes Core and Modem
		$Excel->Run("Get_Stats", "Subsystem", "Others");	# make second graph in Temp w/s
		$Excel->Run("do_Export_Chart_fourth");				# export .png file to cwd
		if ($graph_only ne 'Y') {
			$Excel->Run("$saveashtml_macro_name");			# This db is desired as to be published
		}
		add_to_recipient_list_from_scratch_sheet($PLAssignee_flag, $FAAssignee_flag, $TechLead_flag, $Lead_flag);
		$Excel->Run("Delete_Scratch_and_Temp_Worksheet");	# delete scratch data w/s

	} elsif ($graph_type eq "S_Modem") {		# make one graph
		print "S_Modem.\n";
		$Excel->Run("Make_Scratch_WS_and_Paste");			# put data into scratch data w/s
		$Excel->Run("DeleteRowsModem");						# EXAMPLE DB MANIPULATION - scratch w/s db - deletes all but Core
		$Excel->Run("Get_Stats", "Subsystem", "Modem");		# make second graph in Temp w/s
		$Excel->Run("do_Export_Chart_first");				# export .png file to cwd
		if ($graph_only ne 'Y') {
			$Excel->Run("$saveashtml_macro_name");				# This db is desired as to be published
		}
		add_to_recipient_list_from_scratch_sheet($PLAssignee_flag, $FAAssignee_flag, $TechLead_flag, $Lead_flag);
		$Excel->Run("Delete_Scratch_and_Temp_Worksheet");	# delete scratch data w/s

	} elsif ($graph_type eq "S_Core") {		# make one graph
		print "S_Core.\n";
		$Excel->Run("Make_Scratch_WS_and_Paste");			# put data into scratch data w/s
		$Excel->Run("DeleteRowsCore");						# EXAMPLE DB MANIPULATION - scratch w/s db - deletes all but Core
		$Excel->Run("Get_Stats", "Subsystem", "Core");		# make second graph in Temp w/s
		$Excel->Run("do_Export_Chart_first");				# export .png file to cwd
		if ($graph_only ne 'Y') {
			$Excel->Run("$saveashtml_macro_name");				# This db is desired as to be published
		}
		add_to_recipient_list_from_scratch_sheet($PLAssignee_flag, $FAAssignee_flag, $TechLead_flag, $Lead_flag);
		$Excel->Run("Delete_Scratch_and_Temp_Worksheet");	# delete scratch data w/s

	} elsif ($graph_type eq "S_Others") {		# make one graph
		print "S_Others.\n";
		$Excel->Run("Make_Scratch_WS_and_Paste");			# put data into scratch data w/s
		$Excel->Run("DeleteRowsModemAndCorePositive");						# EXAMPLE DB MANIPULATION - scratch w/s db - deletes all but Core
		$Excel->Run("Get_Stats", "Subsystem", "Others");		# make second graph in Temp w/s
		$Excel->Run("do_Export_Chart_first");				# export .png file to cwd
		if ($graph_only ne 'Y') {
			$Excel->Run("$saveashtml_macro_name");				# This db is desired as to be published
		}
		add_to_recipient_list_from_scratch_sheet($PLAssignee_flag, $FAAssignee_flag, $TechLead_flag, $Lead_flag);
		$Excel->Run("Delete_Scratch_and_Temp_Worksheet");	# delete scratch data w/s
	}

	if ($graph_only eq 'Y') {
		Write_html_Skeleton();
	}

	sleep(1);
	system("move", "$html_file_name", "$sending_file_name"); #rename as embed.html

	EmbedPictureInHtml(); #move pic to current dir, add picture inside html

	$Book->Save();
	return 1;
}

sub Email {
	# accepts one configuration instance
	# calls queryprism
	# calls writeinfotostagingsheet
	# reads data from html, inserts it into email and sends

	my $num_instances = @_;
    if ($num_instances != 1) {
		print "Error when emailing, wrong # of params: @_";
		return 0;
	}

	my %instance_params = %{shift()};
	# the format options -> columns in the final table
	my $columns = shift();

	empty_dir();
#	cleanUpExcel();

	# get info from prism about query, array of hashes
#	print "Query ID - $instance_params{'Query ID'}\n";
#	my @queryresults = QueryPrism($instance_params{'Query ID'});
	# if results exist, print notification
#	if (@queryresults) {
#		print "\nRecieved results from queryprism() for query $instance_params{'Query ID'}\n";
#	}

#	my %sendtohash; # allows for uniqueness of recipients
	# write info to staging sheet then save as html
	if (!WriteInfoToStagingWorksheet(\%instance_params, \@queryresults)) {
		print "Error writing to staging sheet.\n";
	}

	my %temphash; # stores temporary query data

	if (defined $instance_params{'Send To'}) {
		@recipients = split ('\n', $instance_params{'Send To'});
	}
	### Add content to email before sending ###

	addContentToHtml($instance_params{'Text Above Chart'}, $instance_params{'Text Below Chart'});

	push (@recipients, keys %sendtohash);

	if ($instance_params{'Send Once'} eq 'PREVIEW')	{	# If PREVIEW, then overwrite all recipients with only global $preview_members, and blank out all cc's
		$rlist = $preview_members;
		$cc_list = "";
		$instance_params{'Email Subject'} = $instance_params{'Email Subject'} . " - PREVIEW ONLY";
	} else {
		$rlist = join ',', @recipients;
	}

	$rlist =~ s/\s//g;
	$rlist =~ tr{A..Z}{a..z};

	if ($instance_params{'CC'} ne "") {
		$cc_list = "-cc \"$instance_params{'CC'}\"";
	} else {
		$cc_list = "";
	}

	my @files = <*.png>;
	my $size = @files;
	if ($size) {
		$embed_str = ' -embed ';
		foreach $filename (@files) {
			$embed_str = $embed_str . $filename . ',';
		}
	}
	print "rlist is $rlist\ncc_list is $cc_list\n";

	# send email
#	system("$blatlocation \"$sending_file_name\" -server smtphost.qualcomm.com -to \"$rlist\" $cc_list -f \"$instance_params{'Sender'}\"
#		$embed_str -subject \"$instance_params{'Email Subject'}\"");

	return 1;
}

sub ParseParamsAndEmail {
	my $num_instances = @_; # num of arguments, instances of configurations
	if ($num_instances == 0) {
		print "No instances to email. All done.\n";
		return 1;
	}

	print "You have $num_instances config instances to parse\n";

	# interate all instances,
	# check if it is time to send an instance
	# and then send it.
	# If the instance is just a one time send, delete that line in the excel file
	for (my $instance = 0; $instance < $num_instances; $instance++)	{
        for (keys %sendtohash) {
        	delete $sendtohash{$_};
        }
	    $recipients = scalar(@recipients);
    	for ($i = 0; $i<=$recipients; ++$i) {
			$recipients[$i] = undef;
	    }

		if ($_[$instance]{'Send Once'} eq "PREVIEW") {
			$ok_to_process_all = 0;
		}

		if (! $ok_to_process_all) {
			if ($_[$instance]{'Send Once'} eq "PREVIEW") {
				print "Match $_[$instance]{'Configuration Name'}, emailing to preview recipients only\n";
				print "\n=====================================================\n";
				Email($_[$instance]);
				empty_dir();
			}
		} else {
			print "Match $_[$instance]{'Configuration Name'}, emailing\n";
			print "\n=====================================================\n";
			Email($_[$instance]);
			empty_dir();
		}
	}

	return 1;
}

sub Write_html_Skeleton {
	open FILE, ">$html_file_name";
	print FILE "<html>\n";
	print FILE "<body>\n";
	print FILE "<\/body>\n";
	print FILE "<\/html>\n";
	close FILE;
}

sub cleanUpExcel {
	my $i;
	my $j;
	my $k;
	my @deleteary;

	$Excel->{DisplayAlerts} = 0;
	#assume excel is open
	#check for any sheets named Temp, Staging Worksheet, Copy of Staging Worksheet, Temp2, etc.
	my $SheetCount = $Book->Worksheets->Count();
	foreach $i (1..$SheetCount) {
		foreach $j (@TempDeleteSheetNames) {
#			print "Checking for existence of sheet $j\n";
			if ( $Book->Worksheets($i)->{Name} eq $j) {
				print "Deleting $j worksheet\n";
				push (@deleteary, $j);
			}
		}
	}
	foreach $k (@deleteary) {
		$Book->WorkSheets($k)->Delete();
	}
}

sub printTime {
	my @qtime = localtime(time);
	$hour = $qtime[2];
	$min = $qtime[1];
	$sec = $qtime[0];

	if (length($min) eq 1) {
		$min = "0" . $min;
	}
	if (length($sec) eq 1) {
		$sec = "0" . $sec;
	}
	print "$hour:$min:$sec";
}

sub email_name {
	my $in_name = shift;
	my @name_ary = ();
	my @ph_ary = ();
	my $line;
	my $str;

	@name_ary = split(',\s', $in_name);					# last name, first name
	my $search_str = "$name_ary[1] $name_ary[0]";		# first name last name - string to search either hash or in ph - full name is key in both
	if ($exception = $techlead_hash{$search_str}) {		# is name one of those with ambiguous results in ph? Then use file hash instead
		return $exception;
	} else {
		$str = `ph $search_str`;						# otherwise do the ph thing
		@ph_ary = split('\n', $str);
		foreach $line (@ph_ary) {
			if ($line =~ /\s*user_account:\s(\w*)/) {	# and search for line with email acct name in it
				return($1);
			}
		}
		return(0);										# no luck in either.  Hmmmmm......
	}
}

sub add_to_email_list() {
	my $av = $worksheet->Cells(1,$area_column)->{'Value'};
}

sub trythis() {
	my %Areas;
	my $worksheet = $Book->Worksheets("Scratch Staging Worksheet");

	my $threshhold = $LastRow * .94;
	my $total = 0;
	my $i;
	my $j;

	for ($i = 1; $i <= $LastRow; $i++) {
		$myValue = $worksheet->Range("B$i")->{'Value'};
		if (exists $Areas{$myValue}) {
			$Areas{$myValue}++;
		} else {
			$Areas{$myValue} = 1;
		}
	}

	print "Threshold is $threshhold\n================\n";
        foreach $key (sort{$Areas{$b} <=> $Areas{$a}} keys %Areas) {
                if ($total < $threshhold) {
                	$total = $total + $Areas{$key};
	        	print "$Areas{$key}\t$key\n";
                        $keep{$key} = 'True';
                }
	}

	for ($j = 2; $j <= $LastRow; $j++) {
        	$myValue = $worksheet->Range("B$j")->{Value};
		if (! exists $keep{$myValue}) {
                	print "Changing B$j - $myValue : to Others\n";
                	$worksheet->Range("B$j")->{Value} = "Others";
                }
        }

	$Book->Save();
	$Book->Close();
}

sub add_to_recipient_list_from_scratch_sheet() {
	my $PLAssignee_flag = shift;
	my $FAAssignee_flag = shift;
	my $TechLead_flag = shift;
	my $Lead_flag = shift;

	my $i;
	my @NameAry;
	my $PLAssignee_col = 0;
	my $FAAssignee_col = 0;
	my $TechLead_col = 0;
	my $Lead_col = 0;

	my $worksheet = $Book->Worksheets($scratch_staging_worksheet_name);
	my $LastRow_Scratch = $worksheet->UsedRange->Find({What=>"*",
		SearchDirection=>xlPrevious,
		SearchOrder=>xlByRows})->{Row};
	my $LastCol_Scratch = $worksheet->UsedRange->SpecialCells(xlCellTypeLastCell)->{Column};

	for ($i = 1; $i <= $LastCol_Scratch; $i++) {
		if ($worksheet->Cells(1, $i)->{'Value'} eq 'PLAssignee') {
			$PLAssignee_col = $i;
			$Lead_col = $i;
		} elsif ($worksheet->Cells(1, $i)->{'Value'} eq 'FAAssignee') {
			$FAAssignee_col = $i;
		} elsif ($worksheet->Cells(1, $i)->{'Value'} eq 'TechLead') {
			$TechLead_col = $i;
		}
	}

	if ($PLAssignee_flag eq "Y") {
		# if send to PL Assignee is specified, add the list of all of the PL assignees to sendtohash data structure
		for ($i = 2; $i <= $LastRow_Scratch; $i++) {
			@NameAry = split /,/, $worksheet->Cells($i, $PLAssignee_col)->{'Value'};
			foreach $result (@NameAry) {			#find unique names of PLAssignees
				$result =~ s/\s//g;   # remove spaces
				$sendtohash{$result} = "True";
			}
		}
	}

	if ($FAAssignee_flag eq "Y") {
		# if send to FA Assignee is specified, add the list of all of the PL assignees to sendtohash data structure
		for ($i = 2; $i <= $LastRow_Scratch; $i++) {
			@NameAry = split /,/, $worksheet->Cells($i, $FAAssignee_col)->{'Value'};
			foreach $result (@NameAry) {			#find unique names of PLAssignees
				$result =~ s/\s//g;   # remove spaces
				$sendtohash{$result} = "True";
			}
		}
	}

	if ($TechLead_flag eq "Y") {
		# if send to TechLead is specified, add the list of all of the PL assignees to sendtohash data structure
		for ($i = 2; $i <= $LastRow_Scratch; $i++) {
			$result = $worksheet->Cells($i, $TechLead_col)->{'Value'};
			chomp($result);
			$short_name = email_name($result);
			if ($short_name) {							# check just in case email_name returned a nul string
				chomp($short_name);
				$sendtohash{$short_name} = "True";		# adding to hash of email recipients for this Prism query
			}
		}
	}
 	if ($Lead_flag eq "Y") {
		# if sending to leads, add all leads of PLAssignee to list of recipients
		for ($i = 2; $i <= $LastRow_Scratch; $i++) {
			@assignee_ary = split /,/, $worksheet->Cells($i, $Lead_col)->{'Value'}; # actually PLAssignee(s)
            $assigneearynum = @assignee_ary;
			if ($assigneearynum > 0) {		# just in case it was a blank
				foreach $result (@assignee_ary) {				# find unique names of PLAssignees
               		$result =~ s/\s//g;   # remove spaces
					$leadtempstr = $assignee_to_lead{$result};
                    @leadary = split /,/, $leadtempstr;
					$leadarynum = @leadary;
					if ($leadarynum > 0) {
						foreach $lead_str (@leadary) {
							$sendtohash{$lead_str} = "True";
                        }
					}
				}
			}
		}
	}
}