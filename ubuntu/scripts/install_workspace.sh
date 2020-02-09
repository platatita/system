#! /bin/bash -l

source ./install_core.sh
SCRIPTNAME=`basename $0`
trace_start $SCRIPTNAME
# ----------------------------------------------------------------------


trace "install useful tools"
sudo apt -y install nload curl wget htop wipe
sudo apt -y install vim gawk
sudo apt -y install vim-gtk
sleep 1

trace "install git"
sudo apt -y install git-core
sudo apt -y install gitk

trace "install necessary libs"
sudo apt -y install build-essential openssl curl git-core zlib1g zlib1g-dev libssl-dev libyaml-dev libsqlite3-dev sqlite3 libxml2-dev libxslt-dev autoconf libc6-dev ncurses-dev automake libtool bison subversion

trace "install internet speedtest"
sudo apt -y install speedtest-cli
# usage by run: speedtest

# ----------------------------------------------------------------------
trace_end $SCRIPTNAME
