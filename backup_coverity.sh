#!/usr/bin/bash

# Designed to copy original database folder (in it's entirety) in two chronological levels

# real folder names when deployed on thing2 as a cron job from scmbuild
original=/opt/mnt/scm/builds/coverity/data
first_backup=/opt/mnt/scm/builds/coverity/data_backup_1
second_backup=/opt/mnt/scm/builds/coverity/data_backup_2

# test folder names here on scmbuild4
# original=/usr2/jhenk/jim_original
# first_backup=/usr2/jhenk/jim_backup_1
# second_backup=/usr2/jhenk/jim_backup_2

log_file=/opt/mnt/scm/builds/coverity/backup_coverity.log

echo "Log file - backup_coverity" > $log_file
echo "**************************" >> $log_file

#####################################
# basic check to see if database dir is there at all...
if [ ! -d $original ]
then
   perl /usr2/scmbuild/mailme.pl -original_bad
   exit
fi 

export PATH=$PATH:/opt/packages/coverity/prevent-solaris-sparc-3.9.0/bin
cov-stop-gui -d $original

#####################################
# If there, delete oldest copy to make way for more recent copies
if [ -d $second_backup ]
then
   if [ -d $first_backup ]       # if there is no first_backup, then leave second backup alone
   then
      echo "Attempting to delete $second_backup." >> $log_file
      rm -R $second_backup
      if [ -d $second_backup ]
      then
         perl /usr2/scmbuild/mailme.pl -delete_bad
         cov-start-gui -d $original
         exit
      else
         echo "Deletion of $second_backup successful." >> $log_file
      fi
   fi
else
   echo "Did not detect $second_backup No deletion will be attempted." >> $log_file
fi

######################################
# if there, make first (most recent) copy into oldest copy to make room for fresh copy
if [ -d $first_backup ]
then
   echo "Attempting to rename $first_backup to $second_backup." >> $log_file
   mv $first_backup $second_backup
   if [ -d $first_backup ]   # if still there, then the rename was not successful - we have a problem
   then
      perl /usr2/scmbuild/mailme.pl -rename_bad
      cov-start-gui -d $original
      exit
   else
      if [ -d $second_backup ]  # if oldest copy is now there, then we're in good shape.  If not...
      then
         echo "Rename to $second_backup successful." >> $log_file
      else
         perl /usr2/scmbuild/mailme.pl -rename_bad
         cov-start-gui -d $original
         exit
      fi
   fi
else
   echo "Did not detect $first_backup. No rename operation will be attempted." >> $log_file
fi


#######################################
# We already know that the original db dir is in place. (checked above...) So let's make a fresh copy.
echo "Attempting to copy $original to $first_backup - please stand by." >> $log_file
echo "Attempting to copy $original to $first_backup - please stand by."
cp -r $original $first_backup
if [ -d $first_backup ]
then
   echo "Is good."  >> $log_file
   perl /usr2/scmbuild/mailme.pl -groovy
else
   perl /usr2/scmbuild/mailme.pl -copy_bad
fi

cov-start-gui -d $original

