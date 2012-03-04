#! /bin/bash

BROADCASTMASK=192.168.0.255
if [ -n "${1}" ]; then
  BROADCASTMASK=$1
fi

echo IP address to test: $BROADCASTMASK

ping -c 3 $BROADCASTMASK
arp -a

