#! /bin/bash -l

source ./install_core.sh
SCRIPTNAME=`basename $0`
trace_start $SCRIPTNAME
# ----------------------------------------------------------------------


trace "install mandatory packages for rust"
sudo apt -y install build-essential


trace "install rustup"
curl --proto '=https' --tlsv1.2 -sSf https://sh.rustup.rs | sh


# ----------------------------------------------------------------------
trace_end $SCRIPTNAME
