# you have to add this at the end of the .bash_profile file in home directory if it exists otherwise just copy it to the home directory.

if [ -f ~/.bash_aliases ]; then
    . ~/.bash_aliases
fi

if [ -f ~/.bash_custom ]; then
    . ~/.bash_custom
fi

# Added by the Heroku Toolbelt
export PATH="/usr/local/heroku/bin:$PATH"

# brew git
export PATH="/usr/local/bin:$PATH"

# Add the following lines to your ~/.bash_profile:
if [ -f $(brew --prefix)/etc/bash_completion ]; then
     . $(brew --prefix)/etc/bash_completion
fi

# bash history
export HISTCONTROL=ignoredups

[[ -s "$HOME/.rvm/scripts/rvm" ]] && source "$HOME/.rvm/scripts/rvm" # Load RVM into a shell session *as a function*
