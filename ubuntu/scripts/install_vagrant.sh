#! /bin/bash -l

source ./install_core.sh
SCRIPTNAME=`basename $0`
trace_start $SCRIPTNAME
# ----------------------------------------------------------------------


trace "install virtualbox"
sudo apt -y install virtualbox &&\
sudo apt update

trace "download and install vagrant"
cd $HOME/tmp && \
curl -O https://releases.hashicorp.com/vagrant/2.2.7/vagrant_2.2.7_x86_64.deb && \
sudo apt -y install ./vagrant_2.2.7_x86_64.deb && \
cd -


# ----------------------------------------------------------------------
trace_end $SCRIPTNAME
