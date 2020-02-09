#! /bin/bash -l

source ./install_core.sh
SCRIPTNAME=`basename $0`
trace_start $SCRIPTNAME
# ----------------------------------------------------------------------


trace "add sublime apt key"
wget -qO - https://download.sublimetext.com/sublimehq-pub.gpg | sudo apt-key add -

trace "add sublime repository"
sudo apt-add-repository "deb https://download.sublimetext.com/ apt/stable/"

trace "install sublime"
sudo apt -y install sublime-text


# ----------------------------------------------------------------------
trace_end $SCRIPTNAME
