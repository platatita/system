#! /bin/bash
# copy the script into home directory because it is invoked by the alias 'alias sublime='~/.sublime_run.sh' in the .bash_aliases file.
SUBLIMEPATH="/usr/local/bin/sublime_text"

function sublime(){ 
  local tmp_pwd="$(pwd)"
  echo "$SUBLIMEPATH", $tmp_pwd
  "$SUBLIMEPATH" -n $tmp_pwd &
}

sublime

