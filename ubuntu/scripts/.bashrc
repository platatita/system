# you have to add this at the end of the .bashrc file in home directory if it exists otherwise just copy it to the home directory.

if [ -f ~/.bash_aliases ]; then
    . ~/.bash_aliases
fi

if [ -f ~/.bash_custom ]; then
    . ~/.bash_custom
fi