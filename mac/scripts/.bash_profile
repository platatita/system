
# Setting Variable Environment
export PATH=$PATH:/usr/local/mongodb
export PATH=$PATH:/usr/local/mongodb/2.0.3/bin
export PS1='\u@\H:\w$ '

# set terminal colors
export CLICOLOR=1
export LSCOLORS=ExFxCxDxBxegedabagacad

# git
alias gs='git status'
alias gp='git pull'
alias gb='git branch'
alias gl='git log'

# home
alias system='cd ~/sources/git/system'

# anixe
alias anx='cd ~/anx'
alias anxgit='cd ~/anx/sources/git'
alias synw='cd ~/anx/sources/git/synapse_worker'
alias syna='cd ~/anx/sources/git/synapse_aer'
alias resfa='cd ~/anx/sources/git/resfinity_api'
alias bmsw='cd ~/anx/sources/git/bmsweb'
alias bms='cd ~/anx/sources/git/bms'
alias tomaweb='cd ~/anx/sources/git/tomaweb'

# system
alias ll='ls -asl'
alias dtree="find . -type d -print | sed -e 's;[^/]*/;|____;g;s;____|; |;g'"
alias tree="find . -print | sed -e 's;[^/]*/;|____;g;s;____|; |;g'"

OSNAME=`uname`
if [ $OSNAME = 'Linux' ]; then
  alias vbash='vim ~/.bashrc'
  alias sbash='source ~/.bashrc'
else
  alias vbash='vim ~/.bash_profile'
  alias sbash='source ~/.bash_profile'
fi

# Load RVM into a shell session *as a function*
[[ -s "$HOME/.rvm/scripts/rvm" ]] && source "$HOME/.rvm/scripts/rvm"
