#! /bin/bash

if [ -f ./directory_creator_anixe.sh ]; then
  ./directory_creator_anixe.sh
fi

echo "create Downloads folders"
mkdir -pv ~/Downloads/safari
mkdir -pv ~/Downloads/firefox
mkdir -pv ~/Downloads/mailbox
mkdir -pv ~/Downloads/utorrent

echo "create Documents folders"
mkdir -pv ~/Documents/internet
mkdir -pv ~/Documents/games
mkdir -pv ~/Documents/developer/os_hints
mkdir -pv ~/Documents/developer/ios
mkdir -pv ~/Documents/screencapture

echo "create sources folder"
mkdir -pv ~/sources/git
mkdir -pv ~/tmp