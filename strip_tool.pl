#!/usr/bin/perl

######################################################################
#
# Name - strip_tool.pl
# Version - 0.1
# Author - Jim Henk
# Date - 10/29/2012
# Copyright 2012 - Qualcomm Corporation
#
#  Recursive comparison of 2 different labeled sets of the same
#  files, searching for newly introduced #ifdef - #endif sets
#  Input parameters: base, comparison labels, recurse/no recurse option
#
######################################################################

use Getopt::Long;
use File::Basename;
use File::Copy;
use File::Path;        # for deleting directory tree
use Cwd;
use Data::Dumper;
use P4;

my $PROGRAM = fileparse($0);
my $user = $ENV{"USER"};

my $help_opt = 0;
my $firstlabel = "";
my $secondlabel = "";
my $tip_opt = "";
my $cspec = "";
my $outfile = "";
my $verbose = 0;
my $debug = 0;
my $listall_opt = 0;

## Parse command line arguments
Getopt::Long::Configure('pass_through');
GetOptions(
    "help"           => \my $help_opt,
    "firstlabel=s"   => \my $firstlabel,
    "secondlabel=s"  => \my $secondlabel,
    "tip"            => \my $tip_opt,
    "cspec=s"        => \my $cspec,
    "listall"        => \my $listall_opt,
    "revno"          => \my $revno_opt,
    "outfile=s"      => \my $outfile,
    "verbose"        => \my $verbose,
    "debug"          => \my $debug,
);

my $p4 = new P4;
$p4->Connect() or die( "Failed to connect to Perforce Server" );


######################################################################
# Verify command line parameters as syntactically valid
######################################################################
if ($help_opt) {
   usage();
}

if (!$cspec) {
   print "\n*** ERROR - Missing -c (cspec)\n\n";
   usage();
}
if (!$firstlabel) {
   print "\n*** ERROR - Missing -f (first label)\n\n";
   usage();
}
if (!$secondlabel && !$tip_opt) {
   print "\n*** ERROR - Missing -s (second label) or -t (tip)\n\n";
   usage();
}
if ($secondlabel && $tip_opt) {
   print "\n*** ERROR - Choose only one: -s (second label) or -t (tip)\n\n";
   usage();
}

######################################################################

my %file_hash;
my %exceptions_hash;

my $main_client = 'master_' . $cspec;
my $temp_client = $user . '_' . $cspec;
if ($tip_opt) {
	$secondlabel = 'tip';
}

my $scratchdir1 = $ENV{"HOME"} . "/strip_tool_first_dir";
mkdir $scratchdir1;
my $scratchdir2 = $ENV{"HOME"} . "/strip_tool_second_dir";
mkdir $scratchdir2;

my $push_client_name = $p4->GetClient();          # save off what P4CLIENT was at entry

delete_local_client($temp_client);                # clear out the temp client, if it still there for some reason
make_local_client($main_client, $temp_client);    # make new temp client that can deal with strip file dirs

switch_client_root($temp_client, $scratchdir1);   # setup client to load *first* scratch dir
print_client($temp_client);                       # will output edited client IF DEBUG is also selected
msgrouter("Switching to temp client $temp_client\n") if ($verbose);
$p4->SetClient($temp_client);
sync_files($firstlabel, "none");                  # sync files to @first label (without updating P4 metadata)
sync_files($firstlabel, "label");                 # sync files to @first label (without updating P4 metadata)
if (is_folder_empty($scratchdir1)) {              # scratch folder empty if label is empty or non-existant
	msgrouter("*** The label " . $firstlabel . " does not exist in this tree.  Aborting...\n");
	exitrtn();
}
process_dirs ($scratchdir1, length($scratchdir1), $scratchdir1);             # recurse files in first tree,

switch_client_root($temp_client, $scratchdir2);   # setup client to load *second* scratch dir
print_client($temp_client);                       # will output edited client IF DEBUG is also selected
sync_files($firstlabel, "none");                  # sync files to @first label (without updating P4 metadata)
if ($tip_opt) {
   sync_files("", "tip");
} else {
   sync_files($secondlabel, "label");
}
if (is_folder_empty($scratchdir2)) {              # scratch folder empty if label is empty or non-existant
	msgrouter("*** The label " . $secondlabel . " does not exist in this tree.  Aborting...\n");
	exitrtn();
}
process_dirs ($scratchdir2, length($scratchdir2), $scratchdir2);             # recurse files in first tree,

if ($debug) {
   msgrouter("DEBUG - contents of %file_hash:\n", 1);
   $output = Dumper(%file_hash);
   msgrouter($output, 1);
}

if ($listall_opt) {
   do_report_listall();
} else {
   do_report_exceptions();
}

exitrtn();

## ======================================================================
##
## Routine - usage
## input - none
## output - Help Screen dump
## returns - none
##
## ======================================================================
sub usage {
   msgrouter ("*** Usage: $PROGRAM -h\n", 1);
   msgrouter("           $PROGRAM -f firstlabel [-s secondlabel -t] -c cspec <-o outputfile> <-d> <-l> <-r> <-v> <-d>\n", 1);
   msgrouter("\n", 1);
   msgrouter("    -h help (print this screen\n", 1);
   msgrouter("    -f first_label  current setting: $firstlabel\n", 1);
   msgrouter("    -s second_label  current setting: $secondlabel\n", 1);
   msgrouter("    -t (tip of branch used for second label setting: $tip_opt\n", 1);
   msgrouter("    -c specname (template config spec to be used)  current setting: $cspec\n", 1);
   print_master_clients();
   msgrouter("    -l (list all) - current setting: $listall_opt\n", 1);
   msgrouter("    -r (display revision numbers) instead of CLs where ifdef line was introduced (CL is default)  current setting: $revno_opt\n", 1);
   msgrouter("    -o log file name  current setting: $outfile\n", 1);
   msgrouter("    -v (verbose)  current setting: $verbose\n", 1);
   msgrouter("    -d (debug)  current setting: $debug\n", 1);
   
   exit();
}

## ======================================================================
##
## Routine - exitrtn
## input - none
## output - none
## returns - none
##
##  Cleanup rtn - deletes scratch directories and temp P4 client, then exits cleanly
##
## ======================================================================
sub exitrtn {
   delete_local_client($temp_client);                 # clean-up operations
   delete_directory($scratchdir1);
   delete_directory($scratchdir2);
   msgrouter("Switching back to stored client $push_client_name\n") if ($verbose);
   $p4->SetClient($push_client_name);                 # restore P4CLIENT back to what it was before we started
   exit();
}

## ======================================================================
##
## Routine - process_dirs
##  input - root directory name to be processed - RECURSIVE!
##  output - nothing
##  returns - none
##
##  Recursion directory drill-down routine
##  Opens given dir, reads contents 
##       calls process_file for all files
##       calls process_dirs recursively for all dirs
##
## ======================================================================
sub process_dirs {
    my $path = shift;
    my $str_length = shift;
    my $file_hash_key = shift;
    
    # Open the directory.
    opendir (DIR, $path)
        or die "Unable to open $path: $!";

    # Read in the files.
    # You will not generally want to process the '.' and '..' files,
    # so we will use grep() to take them out.
    # See any basic Unix filesystem tutorial for an explanation of them.
    my @files = grep { !/^\.{1,2}$/ } readdir (DIR);

    # Close the directory.
    closedir (DIR);

    # At this point you will have a list of filenames
    #  without full paths ('filename' rather than
    #  '/home/count0/filename', for example)
    # You will probably have a much easier time if you make
    #  sure all of these files include the full path,
    #  so here we will use map() to tack it on.
    #  (note that this could also be chained with the grep
    #   mentioned above, during the readdir() ).
    @files = map { $path . '/' . $_ } @files;

    for (@files) {
        # If the file is a directory
        if (-d $_) {
            # Here is where we recurse
            # This makes a new call to process_files()
            # using a new directory we just found.
            process_dirs ($_, $str_length, $file_hash_key);

        # If it isn't a directory, lets just do some
        # processing on it.
        } else {
           # Do whatever you want here =)
           # A common example might be to rename the file.
           $file_key = substr($_, $str_length);
           # call process_file - parms: absolute path, relative path, directory_root
           process_file($_, $file_key, $file_hash_key);
        }
    }
}

## ======================================================================
##
## Routine - process_file
## input -  full path and file name
##          relative path and path name (full path minus root directory)
##          root directory
## output - %file_hash data structure
## returns - none
##
##  Assumes that all files have been synced successfully into the
##  directory. Reads file, and writes hash record for each #ifdef found.
##
##  %file_hash data structure: hash of hashes of arrays of values
##     major hash key: relative path and file name
##     minor hash key: root directory
##     array: #ifdef string, line number of occurance
##
## ======================================================================
sub process_file {
# call process_file - parms: absolute path, relative path, directory_root

   my $abs_file_path = shift;
   my $relative_path_key = shift;
   my $dir_root_key = shift;
   
   my @ary_of_ifdefs;

   my %ifdef_record = {};
   my $line_number = 0;
   my $true_flag = "0";

   open (INFILE, "< $abs_file_path");
   while (<INFILE>) {
   	  $line_number++;
   	  chomp;
   	  if (/^#ifdef\s*(\w*)/) {                            # build up anonymous hash of #ifdef name and line #
         $ifdef_record->{'line_number'} = $line_number;
         $ifdef_record->{'ifdef_name'} = $_;
         $true_flag = "1";
         push @ary_of_ifdefs, $ifdef_record;              # record for one file in one dir only
   	  }
      $ifdef_record = {};
   }
   close (INFILE);
   
   if ($true_flag) {      
      $file_hash{$relative_path_key}{$dir_root_key} = [@ary_of_ifdefs];
   }
}

## ======================================================================
##
## Routine - do_report_listall
## input - none
## output - Full detail report
## returns - none
##
## Writes a full report on the entire contents of the %file_hash
## data structure by:
##     relative file path and filename - major
##     local root directories for first and second data sets
##     name of the #ifdef declaration, and it's line number in the file
##
## ======================================================================
sub do_report_listall {
	foreach $rel_path_filename (keys(%file_hash)) {
		msgrouter("\n--------------------------------------------\n", 1);
		msgrouter("$rel_path_filename\n", 1);

        for $local_root_dir (keys %{$file_hash{$rel_path_filename}}) {
            msgrouter("========>  root dir - $local_root_dir\n", 1);
            foreach $local_ifdef_array ($file_hash{$rel_path_filename}{$local_root_dir}) {
               foreach $local_ifdef_hash (@{$local_ifdef_array}) {
                  %scratch = %{$local_ifdef_hash};
               	  msgrouter("\t" . $scratch{'ifdef_name'} . " on line " . $scratch{'line_number'} . "\n", 1);
               }
            }
        }
	}
}

## ======================================================================
##
## Routine - do_report_exceptions
## input - none
## output - *local* %exceptions_hash data structure for report          
## returns: none
##
## Reads %file_hash and transforms into local %exceptions_hash data
## structure to be read in order to report on #ifdef declarations
## occurring in second directory tree, but not in the first.
##
## data structure by:
##     name of the #ifdef declaration - major
##     local root directory for first and second data sets - minor
##     key -     relative file path and filename
##     value -   ifdef occurance line number in the file
##
## mini data structure also produced by p4 annotate cmd
##     @annotate_ary - array of hashes
##     $annotate_hash - single hash of:
##         #ifdef name
##         first time (rev # or CL #) ifdef appeared
##         most current time ifdef appeared
##
## ======================================================================
sub do_report_exceptions {
# Read %file_hash, gather data, transform it into %exceptions_hash
	foreach $rel_path_filename (keys(%file_hash)) {
        for $local_root_dir (keys %{$file_hash{$rel_path_filename}}) {
            foreach $local_ifdef_array ($file_hash{$rel_path_filename}{$local_root_dir}) {
               foreach $local_ifdef_hash (@{$local_ifdef_array}) {
# $scratch is a local copy of the ifdef hash that is easier to work with
                  %scratch = %{$local_ifdef_hash};
                  $ifdef_name_key = $scratch{'ifdef_name'};
                  $line_number = $scratch{'line_number'};
# Now that we have the data from the single record in %file_hash,
# transform and re-write record into local %exceptions_hash with
# appropriate control breaks 
                  $exceptions_hash{$ifdef_name_key}{$local_root_dir}{$rel_path_filename} = $line_number;
               }
            }
        }
	}

   if ($debug) {
      msgrouter("DEBUG - contents of %exceptions_hash\n", 1);
      $output = Dumper(%exceptions_hash);
      msgrouter($output, 1);
   }
   
# Now with the exceptions_hash built, onto the reporting step.
# We read the exceptions_hash, and see if there are any #ifdef's that are found in the
# second data set, and NOT in the first data set.  All files containing these #ifdefs
# must each be reported on in detail as *new* #ifdef invocations.

   for $ifdef_name_key (keys(%exceptions_hash)) {                                                 # ifdef_name_key: e.g. #ifdef MYFLAG
      if (exists $exceptions_hash{$ifdef_name_key}{$scratchdir2}) {                               # Does the #ifdef show up in the second data set hash?
         if (! exists $exceptions_hash{$ifdef_name_key}{$scratchdir1}) {                          # Does the #ifdef NOT appear in the first data set hash?
         	for $local_filename_key (keys %{$exceptions_hash{$ifdef_name_key}{$scratchdir2}}) {   # For every file that shows up in the 2nd data set hash, report it
               msgrouter("\n*** New ifdef declation.\n", 1);                                      # print a section headername
               $fullfilename = $scratchdir2 . $local_filename_key;
               @local_array = $p4->Run("where", $fullfilename);                                   # extract depot path for file - neccesary for p4 annotate call
               $local_hash = $local_array[0];                                                     # only one element in ary from p4 where cmd - it's a hash
               $depotFile = $local_hash->{'depotFile'} . "\n";                                    # and there's the //depot path name
               chomp($depotFile);
               if ($revno_opt) {
                  @annotate_ary = $p4->Run("annotate", $depotFile);                               # p4 annotate cmd tells bounds in history of where (rev #s) #ifdef line occurs
                  $column_word = "rev";
               } else {
                  @annotate_ary = $p4->Run("annotate", "-c", $depotFile);                         # p4 annotate cmd also available in chewy CL flavor...
                  $column_word = "CL";
               }

               for $annotate_hash (@annotate_ary) {
                  if ($annotate_hash->{'data'} =~ /($ifdef_name_key)/) {                          # report full P4 pathname, #ifdef name, line #, & 1st place in history it occurs
                  	 msgrouter("$depotFile - $1 ", 1);
                  	 msgrouter("at line " . $exceptions_hash{$ifdef_name_key}{$scratchdir2}{$local_filename_key}, 1);
                  	 msgrouter(" introduced at $column_word $annotate_hash->{'lower'}\n", 1);
                  } 
               } 
         	}
         }
      }
   }
   msgrouter("\n", 1);
}

## ======================================================================
##
## routine: switch_client_root <>
## input - clientspec name
##         Root: directory name to be substituted in to clientspec
##         P4Perl data structure containing client spec to be edited
## output - Nothing      
## returns: none
##
## Swaps Root: value in clientspec in Perforce.  This is not a concern
## here, as in this script, the function is only intended to edit
## a temporary clientspec that is disposed of at the end of the script's run.
##
## ======================================================================
sub switch_client_root() {
	my $local_temp_client = shift;
	my $local_root_dir = shift;
    my $local_hash;
    
    msgrouter("Switching client root to $local_root_dir\n") if ($verbose);
    $local_hash = $p4 -> FetchClient($local_temp_client);
    $local_hash->{'Root'} = $local_root_dir;
    $p4 ->SaveClient($local_hash);
}

## ======================================================================
##
## routine: print_client <>
## input - clientspec name
## output - edited client ONLY IF DEBUG IS ENABLED      
## returns: none
##
## Debugging tool only for producing and simple printing of
## clientspec hash from P4Perl call if '-d' is selectd on command line.
##
## ======================================================================
sub print_client() {
	$local_temp_client = shift;
	
   if ($debug) {
      msgrouter("DEBUG - Edited Client:\n", 1);
      $local_hash = $p4->Run("client", "-o", $local_temp_client);
      $output = Dumper($local_hash);
      msgrouter($output, 1);
   }
}

## ======================================================================
##
## routine: make_local_client <>
## input - clientspec name
##         name of client to produce from clientspec name template
## output - Nothing      
## returns: none
##
## ======================================================================
sub make_local_client() {
	my $temp_main_client = shift;
	my $temp_local_client = shift;

    msgrouter("Creating scratch client - $temp_client\n") if ($verbose);
    $p4->Run("client", "-s", "-t", "$temp_main_client", "$temp_local_client");
}

## ======================================================================
##
## routine: delete_local_client <>
## input - clientspec name to be deleted from Perforce
## output - Nothing      
## returns: none
##
## ======================================================================
sub delete_local_client() {
	my $temp_client = shift;

    msgrouter("Deleting client $temp_client\n") if ($verbose);
    $p4->Run("client", "-d", "$temp_client");
}

## ======================================================================
##
## routine: print_master_clients
## input - none
## output - Nothing      
## returns: none
##
## Called by usage screen help routine, prints clients that begin
## with the prefix 'master_'
##
## ======================================================================
sub print_master_clients() {
    msgrouter("\tPotential clients are:\n", 1);
    $clients = $p4->Run("clients");
    @ary_of_clients = @$clients;
    for $temp_hash (@ary_of_clients) {
       $temp_client = $temp_hash->{'client'};
       if ($temp_client =~ (/^master_(\w*)/)) {
       	  msgrouter("\t\t$1\n", 1);
       }
    }
}

## ======================================================================
##
## routine: delete_directory
## input - name of directory tree to wipe away unconditionally
## output - Nothing      
## returns: none
##
## ======================================================================
sub delete_directory() {
	my $dir_to_delete = shift;
	
    msgrouter("\Removing directory: $dir_to_delete\n") if ($verbose);
    if (-d $dir_to_delete) {
       rmtree($dir_to_delete);
    }
}

## ======================================================================
##
## routine: sync_files
## input - label to sync to
##         flag to select how sync command line parameters are to
##         be constructed.
##    Type possibilities: literal label
##                        none - for clearing have list
##                        tip - sync from tip
## output - Nothing      
## returns: none
##
## Builds sync command line, (call to P4Perl module)
## and executes sync. The '-p' parameter allows for syncing
## that does *not* update the p4.have meta-data
##
## ======================================================================
sub sync_files() {
    my $label = shift;
    my $sync_type = shift;

    my $sync_parm;
    
    if ($sync_type eq "label") {
    	$sync_parm = "@" . $label;
    } elsif ($sync_type eq "tip") {
    	$sync_parm = '#head';
    } elsif ($sync_type eq "none") {
    	$sync_parm = '#none';
    } else {
    	msgrouter("sync_files call with type declared as $sync_type is invalid.");
    	exit();
    }
	    
    msgrouter("Syncing to $sync_parm\n") if ($verbose);
    msgrouter("sync command is: \$p4\-\>Run(\"sync\", \"-p\", \"$sync_parm\")\n") if ($verbose);
    $output = $p4->Run("sync", "-p", "$sync_parm");
    if ($debug) {
       msgrouter("DEBUG - sync output\n");
       $suboutput = Dumper($output);
       msgrouter($suboutput);
    }
}

## ======================================================================
##
## Routine - is_folder_empty
## input - none
## output - none
## returns - none
##
##  Utility rtn - merely return true if passed directory has no contents besides '.' and '..'
##
## ======================================================================
sub is_folder_empty {
   my $dirname = shift;
   opendir(my $dh, $dirname);
   return scalar(grep {$_ ne "." && $_ ne ".." } readdir($dh)) == 0;
}

## ======================================================================
##
## routine: msgrouter <>
## input - output logging line
##         flag to select how sync command line parameters are to
##         be constructed.
##    Type possibilities: literal label
##                        none - for clearing have list
##                        tip - sync from tip
## output - Nothing      
## returns: none
##
## ======================================================================
sub msgrouter() {
    my $msg     = shift;
    my $nostamp = shift;

    if ($outfile) {
        if ( -e $outfile ) {
            open LOG, ">>$outfile"
              or die "[Error] failed to open $outfile $!";
        }
        else {
            open LOG, ">$outfile"
              or die "[Error] failed to open $outfile $!";
        }
        print LOG "[" . scalar localtime() . "] " unless $nostamp;
        print LOG "$msg";
        close LOG;
    } else {
        print "[" . scalar localtime() . "] " unless $nostamp;
        print "$msg";
    }
}
