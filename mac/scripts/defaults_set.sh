#! /bin/bash

if [ -d ~/Dropbox/system/mac/screencapture ]; then
	defaults write com.apple.screencapture location ~/Dropbox/system/mac/screencapture
	echo "set screen capture location to: ~/Dropbox/system/mac/screencapture"
elif [ -d ~/SkyDrive/system/mac/screencapture ]; then
	defaults write com.apple.screencapture location ~/SkyDrive/system/mac/screencapture
	echo "set screen capture location to: ~/SkyDrive/system/mac/screencapture"
else
	defaults write com.apple.screencapture location ~/Documents/screencapture
	echo "set screen capture location to: ~/Documents/screencapture "
fi

killall SystemUIServer

echo "set display full location in finder window"
defaults write com.apple.finder _FXShowPosixPathInTitle -bool YES

echo "force finder to show hidden files"
defaults write com.apple.finder AppleShowAllFiles TRUE
osascript -e 'tell app "Finder" to quit'

echo "show unsupported network volumes"
defaults write com.apple.systempreferences TMShowUnsupportedNetworkVolumes 1

