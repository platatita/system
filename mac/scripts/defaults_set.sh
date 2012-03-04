#! /bin/bash

# set screen capture location
defaults write com.apple.screencapture location ~/Documents/screencapture
killall SystemUIServer

# set display full location in finder window
defaults write com.apple.finder _FXShowPosixPathInTitle -bool YES
# force finder to show hidden files
defaults write com.apple.finder AppleShowAllFiles TRUE
osascript -e 'tell app "Finder" to quit'

# show unsupported network volumes
defaults write com.apple.systempreferences TMShowUnsupportedNetworkVolumes 1

