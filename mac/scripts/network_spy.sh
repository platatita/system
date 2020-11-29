#! /bin/bash

BROADCASTMASK=192.168.1.255
if [ -n "${1}" ]; then
  BROADCASTMASK=$1
fi

echo "-------- ping: IP address to test: $BROADCASTMASK --------"
# ping -- send ICMP ECHO_REQUEST packets to network hosts
# 	-c: 	count, tries 3 times
ping -c 3 $BROADCASTMASK


echo "-------- arp --------"
# arp -- address resolution display and control
# 	-a: 	The program displays or deletes all of the current ARP entries.
arp -a

