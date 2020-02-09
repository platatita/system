#! /bin/bash -l

source ./install_core.sh
SCRIPTNAME=`basename $0`
trace_start $SCRIPTNAME
# ----------------------------------------------------------------------

# Install dirvers RTL8814AU for ASUS USB-AC68

tracd "Install dkms (Dynamic Kernel Module Support)"
sudo apt -y install dkms

trace "Clone the repo with the branch 5.6.4.2"
cd $HOME/tmp && \
git clone -b v5.6.4.2 https://github.com/aircrack-ng/rtl8812au.git && \
cd rtl8812au

trace "Install rtl8812au driver"
sudo ./dkms-install.sh

# after sucessful installation of above steps the internet connection should be possible
# and then the website should be visited: https://www.asus.com/us/Networking/USB-AC68/HelpDesk_Download/

# ----------------------------------------------------------------------
trace_end $SCRIPTNAME