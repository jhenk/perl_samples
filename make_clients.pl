#!/usr/bin/perl
######################################################################
##
## Name - make_clients.pl
my $versno = "0.9";
## Author - Jim Henk
## Date - 10/15/2013
## Copyright 2013 - Qualcomm Corporation
##
##   Creates Perforce client(s) for QCT Developmemnt staff of
##   PLs and CPLs currently stored in Perforce.
##
##   Clients created include [parameter]_cute, [parameter]_mute, and [parameter]_sute
##   for allowed PLs, (currently NI.1.2, and NI.1.3)
##
##   One client created for new structure in CPLs at NI.4.0 and greater
##
##   Known issue: supplied root dir (-r option) not checked for validity
##
######################################################################

use Data::Dumper;
use Getopt::Long;
use Cwd;
use File::Basename;

my $commercial_file = 'commercial.txt';
my $source_file = 'source.txt';
my $spec_file = 'spec.txt';
my $clientexam_file = 'clientexam.txt';
my $type;
my $post;
my $PROGRAM = fileparse($0);

## Parse command line arguments
Getopt::Long::Configure('pass_through');
GetOptions(
"help"              => \my $help_opt,
"info"              => \my $info_opt,
"workspace=s"       => \my $workspace_opt,
"branch=s"          => \my $branch_opt,
"rootdir=s"         => \my $rootdir_opt,
"cute"              => \my $cute_opt,
"sute"              => \my $sute_opt,
"mute"              => \my $mute_opt,
"all"               => \my $all_opt,
"post=s"            => \my $post_opt,
"nosync"            => \my $nosync_opt,
"yes"               => \my $yes_opt,
"verbose"           => \my $verbose_opt,
"debug"             => \my $debug_opt,
);

if ($info_opt || $help_opt) {
	if ($info_opt) {
		print "*** $PROGRAM - Version: $versno\n\n";
	}
	if ($help_opt) {
		usage();
	}
    exitrtn();
}

if (!$branch_opt) {
	print "*** Missing -b option - quitting\n";
	usage();
	exitrtn();
}

my $branch_type = is_it_PL_or_CPL($branch_opt);			# PL if specific values, CPL all else
my $found_in_comm = found_in_commercial($branch_opt);	# //commercial/* search
my $old_or_new = is_it_at_least_NI_4_0($branch_opt);	# 0 is old, 1 is new
my $found_in_src = found_in_source($branch_opt);		# //source/* search

if ($debug_opt) {
	print "Evaluated $branch_opt\n";
	print "This makes it: (old or new) $old_or_new (PL or CPL) $branch_type\n";
	print "               (found in comm) $found_in_comm (found in src) $found_in_src\n";
}

if ($branch_type eq 'CPL' && $old_or_new eq 'new') {
	if ($cute_opt || $mute_opt || $sute_opt || $all_opt) {
		print "*** ERROR: -c, -s, -m, or -a invalid for Package Warehouse CPLs\n";
		usage();
		exitrtn();
	}
} elsif (!$cute_opt && !$mute_opt && !$sute_opt && !$all_opt) {
	print "*** ERROR: Must specify either -c, -s, -m, or -a - quitting\n";
	usage();
	exitrtn();
} elsif (($cute_opt && $mute_opt) || ($cute_opt && $sute_opt) || ($cute_opt && $all_opt) ||
($sute_opt && $mute_opt) || ($sute_opt && $all_opt) ||
($mute_opt && $all_opt)) {
	print "*** ERROR: Combination not allowed.  Must specify either -c, -s, -m or -a - quitting\n";
	usage();
	exitrtn();
}

if ($all_opt) {
	$cute_opt = 1;
	$sute_opt = 1;
	$mute_opt = 1;
}

if (!$workspace_opt) {
	$workspace_opt = $ENV{"USERNAME"} . '_' . $branch_opt . $post;
}

if ($rootdir_opt =~ /\s/) {
	print "*** ERROR - Space(s) detected in proposed root directory: $rootdir_opt - not allowed - quitting.\n";
	usage();
	exitrtn();
}

$rootdir_tmp = $rootdir_opt;
$workspace_tmp = $workspace_opt;
if ($debug_opt) {
	print "Opening state:\n";
	print "\trootdir_tmp:   $rootdir_tmp\n";
	print "\tworkspace_tmp: $workspace_tmp\n";
	print "\tcute: $cute_opt    mute: $mute_opt    sute: $sute_opt\n";
}

if ($branch_type eq 'PL') {
	if ($branch_opt eq 'MPSS.NI.1.2' || $branch_opt eq 'MPSS.NI.1.3') {
		if ($cute_opt) {
			$rootdir_opt = process_rootdir($rootdir_tmp, $workspace_tmp, 'cute');
			$workspace_opt = process_workspace($workspace_tmp, 'cute');
			make_and_sync_local_client('firstcase', 'cute');
        }
		if ($mute_opt) {
			$rootdir_opt = process_rootdir($rootdir_tmp, $workspace_tmp, 'mute');
			$workspace_opt = process_workspace($workspace_tmp, 'mute');
			make_and_sync_local_client('firstcase', 'mute');
		}
		if ($sute_opt) {
			$rootdir_opt = process_rootdir($rootdir_tmp, $workspace_tmp, 'sute');
			$workspace_opt = process_workspace($workspace_tmp, 'sute');
			make_and_sync_local_client('firstcase', 'sute');
		}
	} else {		# invalid PL
		if (!$found_in_comm) {
			print "*** ERROR - $branch_opt not found in //commercial/* - quitting\n";
			exitrtn();
		} elsif (!$found_in_src) {
			print "*** ERROR - UT files are not ready yet for $branch_opt, please contact integration team - quitting\n"; # not found in //source
			exitrtn();
		} else {
			print "*** ERROR - $branch_opt not eligible for client creation - quitting\n";
			exitrtn();
		}
	}
} elsif ($branch_type eq 'CPL') {
	if (!$found_in_comm) {
		print "*** ERROR - $branch_opt not found in //commercial/* - quitting\n";
		exitrtn();
	}
    
	if (!$found_in_src) {
		print "*** ERROR - UT files are not ready yet for $branch_opt, please contact integration team - quitting\n"; # not found in //source
		exitrtn();
	}
    
	if ($old_or_new eq 'old') {
		if ($cute_opt) {
			$rootdir_opt = process_rootdir($rootdir_tmp, $workspace_tmp, 'cute');
			$workspace_opt = process_workspace($workspace_tmp, 'cute');
			make_and_sync_local_client('thirdcase', 'cute');
		}
		if ($mute_opt) {
			$rootdir_opt = process_rootdir($rootdir_tmp, $workspace_tmp, 'mute');
			$workspace_opt = process_workspace($workspace_tmp, 'mute');
			make_and_sync_local_client('thirdcase', 'mute');
            
		}
		if ($sute_opt) {
			$rootdir_opt = process_rootdir($rootdir_tmp, $workspace_tmp, 'sute');
			$workspace_opt = process_workspace($workspace_tmp, 'sute');
			make_and_sync_local_client('thirdcase', 'sute');
		}
	} else {	# it's 'new'
		$rootdir_opt = process_rootdir($rootdir_tmp, $workspace_tmp);
		$workspace_opt = process_workspace($workspace_tmp);
		make_and_sync_local_client('fourthcase', '');
	}
} else {
	print "*** Error - $branch_opt is neither a PL, nor a CPL - rejected - quitting\n";
	usage();
	exitrtn();
}

exitrtn();

sub process_rootdir {
	$local_rootdir = shift;
	$local_workspace = shift;
	$local_ext = shift;
	
	if ($debug_opt) {
		print "local_rootdir at beginning of rtn is $local_rootdir\n";
		print "local_ext is $local_ext\n";
	}
	
	if (!$local_rootdir) {
		my $pwd = getcwd();
		$local_rootdir = $pwd . "\\" . $local_workspace;
		$local_rootdir =~ s{/}{\\}g;
	}
    
	if ($local_ext) {
		$local_rootdir = $local_rootdir . '_' . $local_ext;
	}
	
	if ($post_opt) {
		$local_rootdir = $local_rootdir . '_' . $post_opt;
	}
    
	if ($debug_opt) {
		print "local_rootdir at end of rtn is $local_rootdir\n";
	}
	
	return $local_rootdir;
}

sub process_workspace {
	$local_workspace = shift;
	$local_ext = shift;
	
	if ($debug_opt) {
		print "local_workspace at beginning of rtn is $local_workspace\n";
	}
	
	if ($local_ext) {
		$local_workspace = $local_workspace . '_' . $local_ext;
	}
	
	if ($post_opt) {
		$local_workspace = $local_workspace . '_' . $post_opt;
	}
    
	if ($debug_opt) {
		print "local_workspace at end of rtn is $local_workspace\n";
	}
    
	return $local_workspace;
}

## ======================================================================
##
## routine: is_it_PL_or_CPL <>
## input - None - looks at command-line parameter $branch_opt
##                (e.g - MPSS.NI.4.0.1.7 - Segments are '.' delimited)
## output - None
## returns - 'PL' if specific values
##           'CPL' all else
##
## ======================================================================
sub is_it_PL_or_CPL() {
	my $cmpstring = lc($branch_opt);
	if (($cmpstring eq 'mpss.ni.1.2') || ($cmpstring eq 'mpss.ni.1.3')) {
		$answer = 'PL';				# PL
	} else {
		$answer = 'CPL';			# CPL
	}
    
	return $answer;
}

## ======================================================================
##
## routine: make_and_sync_local_client <>
## input - var to pass to View constuct rtn: 'firstcase', 'secondcase', 'thirdcase', 'fourthcase'
##         var to pass to View constuct rtn: 'cute', 'sute', 'mute'
## output - created client, which is synced to tip unless user selected nosync on command line
## returns: None
##
##  Creates one Perforce client - constructs scratch file, imports it into Perforce,
##  (if not deselected) - creates root dir of new client, and syncs to the tip.
##
## ======================================================================
sub make_and_sync_local_client() {
	my $make_local_client_case = shift;
	my $client_extension = shift;
	
	#  Before bothering to write clientspec out to a file,
	#  let's see if this client already exists in Perforce.
	#  If it already does, then this may all be a dreadful mistake...
	#  ...unless, of course, the user has put a '-y' in the
	#  command line.  If so, then he's on his own.  Wow.  Brave...
	system("p4 clients > $clientexam_file");				# Write a file with the names of all the P4 clients in it.
	open (EXAMFILE, $clientexam_file) or die "Could not open $clientexam_file.\n";
	while (<EXAMFILE>) {										# Is ours in the file?
		if ((/^Client\s$workspace_opt\s/i) && (!$yes_opt)) {		# Warn the user, unless he's opted not to be warned.
			print "*** WARNING - $workspace_opt already exists in Perforce - use -y parameter to overwrite\n";
			close EXAMFILE;
			return;
		}
	}
	close EXAMFILE;
    
	print "Processing client - $workspace_opt\n";
	
	# Set up the contents of the client spec, ultimately in the scalar $spec_text
	my $user = $ENV{"USERNAME"};
	my $client_line = "Client: $workspace_opt";
	my $comp_name = $ENV{COMPUTERNAME};
	my $local_view_scalar = fill_out_view_scalar($make_local_client_case, $client_extension);
    
    $spec_text = "$client_line
Update: 2013/04/02 11:12:19
Access: 2013/04/24 13:04:10
Owner:  $user
Host:   $comp_name
Description:
    Created by $user.
Root:  $rootdir_opt
Options:        noallwrite noclobber nocompress unlocked nomodtime normdir
SubmitOptions:  revertunchanged
LineEnd:        local
View:
    $local_view_scalar\n";
    
	# Write it out to a file, as that's the pipe we'll use to feed the 'p4 client -i' command
	open (SPECFILE, ">$spec_file") or die "Could not open file $specfile.\n";
	print SPECFILE $spec_text;
	close SPECFILE;
    
	# Do it.
	system("p4 client -i  <$spec_file");
    
	if (! $nosync_opt) {
		if (! -e $rootdir_opt) {
			print "   Creating root dir - $rootdir_opt...\n";
			mkdir ($rootdir_opt,0777);
		}
        
		$ENV{"PWD"} = $rootdir_opt;
		print "   Syncing...\n";
        chdir "$rootdir_opt";
		system("p4 -c $workspace_opt sync ... > nul");
	} else {
		print "   No client sync.\n";
	}
    
	print "\n";
}


## ======================================================================
##
## routine: fill_out_view_scalar
## input - Major test case name, and clientname extension, if applicable
## returns - Scalar with entire filled out 'View:' statement for Perforce client
##
##  Calls unique correct 'View' scalar construction routine (out of 10) based on input vars
##
## ======================================================================
sub fill_out_view_scalar() {
	my $local_case = shift;
	my $local_extension = shift;
	
	if ($local_case eq 'firstcase') {
		if ($local_extension eq 'cute') {
			return fill_out_view_scalar_old_PL_firstcase_cute();
		} elsif ($local_extension eq 'mute') {
			return fill_out_view_scalar_old_PL_firstcase_mute();
		} elsif ($local_extension eq 'sute') {
			return fill_out_view_scalar_old_PL_firstcase_sute();
		} else {
			print "*** ERROR in call to fill_out_view_scalar - $local_case and $local_extension invalid combination - contact programmer\n";
			exitrtn();
		}
	} elsif ($local_case eq 'thirdcase') {
		if ($local_extension eq 'cute') {
			return fill_out_view_scalar_old_CPL_cute();
		} elsif ($local_extension eq 'mute') {
			return fill_out_view_scalar_old_CPL_mute();
		} elsif ($local_extension eq 'sute') {
			return fill_out_view_scalar_old_CPL_sute();
		} else {
			print "*** ERROR in call to fill_out_view_scalar - $local_case and $local_extension invalid combination - contact programmer\n";
			exitrtn();
		}
	} elsif ($local_case eq 'fourthcase') {
		return fill_out_view_scalar_new_CPL();
	}
}

## ======================================================================
##
## routine: fill_out_view_scalar_old_PL_firstcase_cute <>    case #1 for '_cute'
## input - None
## output - None
## returns: proper view scalar
##
## ======================================================================
sub fill_out_view_scalar_old_PL_firstcase_cute() {
	return "	//source/qcom/qct/modem/mmode/cm/rel/nikel.1.1/modem_proc/... //$workspace_opt/modem_proc/...\n";
}

## ======================================================================
##
## routine: fill_out_view_scalar_old_PL_firstcase_mute <>    case #1 for '_mute'
## input - None
## output - None
## returns: proper view scalar
##
## ======================================================================
sub fill_out_view_scalar_old_PL_firstcase_mute() {
	return "	//source/qcom/qct/modem/mmode/cm/rel/nikel.1.1/api/... //$workspace_opt/api/...\n";
}

## ======================================================================
##
## routine: fill_out_view_scalar_old_PL_firstcase_sute <>    case #1 for '_sute'
## input - None
## output - None
## returns: proper view scalar
##
## ======================================================================
sub fill_out_view_scalar_old_PL_firstcase_sute() {
	return "	//source/qcom/qct/modem/mmode/cm/rel/nikel.1.1/doc/... //$workspace_opt/doc/...\n";
}

## ======================================================================
##
## routine: fill_out_view_scalar_old_CPL_cute <>   case #3 for '_cute'
## input - None
## output - None
## returns: proper view scalar
##
## ======================================================================
sub fill_out_view_scalar_old_CPL_cute() {
	return "	//source/qcom/qct/modem/mmode/cm/rel/nikel.1.1/inc/... //$workspace_opt/inc/...\n";
}

## ======================================================================
##
## routine: fill_out_view_scalar_old_CPL_mute <>   case #3 for '_mute'
## input - None
## output - None
## returns: proper view scalar
##
## ======================================================================
sub fill_out_view_scalar_old_CPL_mute() {
	return "	//source/qcom/qct/modem/mmode/cm/rel/nikel.1.1/modem_api/... //$workspace_opt/modem_api/...\n";

}

## ======================================================================
##
## routine: fill_out_view_scalar_old_CPL_sute <>   case #3 for '_sute'
## input - None
## output - None
## returns: proper view scalar
##
## ======================================================================
sub fill_out_view_scalar_old_CPL_sute() {
	return "	//source/qcom/qct/modem/mmode/cm/rel/nikel.1.1/src/... //$workspace_opt/src/...\n";
}

## ======================================================================
##
## routine: fill_out_view_scalar_new_CPL <>    case #4 for CPLs greater than NI.4.0, DI, and others
## input - None
## output - None
## returns: proper view scalar
##
## ======================================================================
sub fill_out_view_scalar_new_CPL() {
	return "	//source/qcom/qct/modem/mmode/cm/rel/nikel.1.1/test/... //$workspace_opt/test/...\n";
}

## ======================================================================
##
## routine: is_it_at_least_NI_4_0 <>
## input - string to be parsed
## output - None
## returns: 'new' if 'greater', 'old' if not
##
## Determines if parameter string is either (less than NI.4.0), or
## (greater than NI.4.0, DI, TR, or BO)
##
## ======================================================================
sub is_it_at_least_NI_4_0() {
	my $cmpstring = shift;
	my ($prefix, $version) = $cmpstring =~ /^MPSS[_\.](\w*)[_\.](.*)/i;
    
	if (lc($prefix) eq 'ni') {
		if ($version lt '4.0') {
			$answer = 'old';		# no, it's lower than NI_4.0
		} else {
			$answer = 'new';		# yes, it's at least NI_4.0
		}
	} else {
		$answer = 'new';			# yes, it's at least NI_4.0
	}
    
	return $answer;
}

## ======================================================================
##
## routine: found_in_commercial <>
## input - branch name to search for
## output -
## returns: 0 if found, 1 if not found
##
## Searches //commercial for existence of named branch. (directory name)
##
## ======================================================================
sub found_in_commercial() {
	my $return_value = 0;
    
	if ($verbose_opt) {
		print "Searching in //commercial/*\n";
	}
    
	@viewdirs = system("p4 dirs //commercial/* > $commercial_file");
	open (COMMERCIALFILE, $commercial_file) or die "Could not open file $commercial_file.\n";
	while (<COMMERCIALFILE>) {
		if (/\/\/commercial\/$branch_opt$/i) {
			if ($verbose_opt) {
				print "Found it - $_\n";
			}
			$return_value = 1;
		}
	}
	close COMMERCIALFILE;
    
	return $return_value;
}

## ======================================================================
##
## routine: found_in_source <>
## input - branch name to search for
## output -
## returns: 0 if found, 1 if not found
##
## Searches one of the two //source structures for existence of named branch. (directory name)
##          ~test/cm/rel if less than NI.4.0, ('old')
##          ~test/mmode/rel if NI.4.0 or greater ('new')
##
## ======================================================================
sub found_in_source() {
	my $local_branch_opt = shift;
	my @local_array;
	my $num_of_elements;
	my $i;
	my $dir;
	my $value;
	my @fields;
	my $ext;
	my $return_value = 0;
	my $index;
    
	my @dir_ary = ("//source/qcom/qct/modem/mmode/test/cm/rel/*", "//source/qcom/qct/modem/mmode/test/mmode/rel/*");
	my $index;
	if ($old_or_new eq 'old') {
		$index = 0;
	} else {
		$index = 1;
	}
    
	$dir = $dir_ary[$index];							# dir prefix for old or new structure - just needs branch-name
	
	if ($verbose_opt) {
		print "Searching in $dir\n";
	}
    
	system("p4 dirs $dir > $source_file");
	chop($dir);												# lose the trailing '*'
	chop($dir);												# lose the trailing '/'
	open (SOURCEFILE, $source_file) or die "Could not open file $source_file.\n";
	while (<SOURCEFILE>) {
		if ($debug_opt) {
			print "SOURCE - branch_opt is " . lc($branch_opt) . " and loop var is $_***\n";
		}
		if (/$dir\/$branch_opt$/i) {
			if ($verbose_opt) {
				print "Found it - $_\n";
			}
			$return_value = 1;
		}
	}
	close SOURCEFILE;
    
	return $return_value;
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
    print "           $PROGRAM [-b branchname] [-c, -s, -m, -a] <-n> <-w workspacename> <-r rootdir> <-p post-text> <-y> <-v> <-d>\n";
    print "\n";
    print "    -h help - print this screen and quits   Current setting - $help_opt\n";
    print "    -i info - prints script version number and quits   Current setting - $info_opt\n";
    print "    -b branch name to be searched for   Current setting - $branch_opt\n";
    print "    -w desired workspace name   Current setting - $workspace_opt\n";
    print "    -r desired root directory  ** Embedded spaces are not allowed\n";
    print "       Current setting - $rootdir_opt\n";
    print "\n";
    print "        In the following set of four, choose only one: -c, -s, -m, or -a (not allowed in PW-based CPLs)\n";
    print "    -c creates 'cute' workspace   Current setting - $cute_opt\n";
    print "    -s creates 'sute' workspace   Current setting - $sute_opt\n";
    print "    -m creates 'mute' workspace   Current setting - $mute_opt\n";
    print "    -a creates all 3 workspaces   Current setting - $all_opt\n";
    print "    -n no-sync (default is to sync client(s) to tip)  Current setting - $nosync_opt\n";
    print "\n";
    print "    -p post text for clientname (e.g. CR98765)   Current setting - $post_opt\n";
    print "    -y override warnings - CAUTION!   Current setting - $yes_opt\n";
    print "    -v verbose (allows on-screen display of script status)   Current setting - $verbose_opt\n";
    print "    -d debug mode\n";
    print "\n";
}

## ======================================================================
##
## Routine - exitrtn
## input - None
## output - None
## returns - None
##
##    Mop-up operations
##
## ======================================================================
sub exitrtn {
	if (-e $commercial_file) {
		if ($verbose_opt) {
			print "Deleting scratch commercial text file\n";
		}
		unlink($commercial_file);
	}
	if (-e $source_file) {
		if ($verbose_opt) {
			print "Deleting scratch source text file\n";
		}
		unlink($source_file);
	}
	if (-e $spec_file) {
		if ($verbose_opt) {
			print "Deleting scratch spec text file\n";
		}
		unlink($spec_file);
	}
	if (-e $clientexam_file) {
		if ($verbose_opt) {
			print "Deleting scratch client exam text file\n";
		}
		unlink($clientexam_file);
	}
	exit();
}


