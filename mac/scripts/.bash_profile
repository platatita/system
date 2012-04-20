# Setting Variable Environment
# add mongo to system PATH
export PATH=$PATH:/usr/local/mongodb
export PATH=$PATH:/usr/local/mongodb/2.0.3/bin
# dispaly in the terminal the username@hostname:current directory path.
export PS1='\u@\H:\w$ '

# set terminal colors
export CLICOLOR=1
export LSCOLORS=ExFxCxDxBxegedabagacad


# Load RVM into a shell session *as a function*
[[ -s "$HOME/.rvm/scripts/rvm" ]] && source "$HOME/.rvm/scripts/rvm"
