#! /bin/bash

# copy all files within current folder to the Sublime Text 3 location

current_user=`whoami`

cp -v ./*.sublime-build "/Users/$current_user/Library/Application Support/Sublime Text 3/Packages/User/"