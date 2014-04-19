# you have to add this at the end of the .bash_profile file in home directory if it exists otherwise just copy it to the home directory.

if [ -f ~/.bash_aliases ]; then
    . ~/.bash_aliases
fi

if [ -f ~/.bash_custom ]; then
    . ~/.bash_custom
fi

# macports bash-completion link: http://trac.macports.org/wiki/howto/bash-completion
if [ -f /opt/local/etc/profile.d/bash_completion.sh ]; then
  source /opt/local/etc/profile.d/bash_completion.sh
fi

# Added by the Heroku Toolbelt
export PATH="/usr/local/heroku/bin:$PATH"

# bash history
export HISTCONTROL=ignoredups

 
export PATH=/opt/local/bin:/opt/local/sbin:$PATH
# Finished adapting your PATH environment variable for use with MacPorts.

# Load RVM into a shell session *as a function*
[[ -s "$HOME/.rvm/scripts/rvm" ]] && source "$HOME/.rvm/scripts/rvm" 
