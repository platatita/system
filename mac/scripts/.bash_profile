# The issue is that Bash sources from a different file based on what kind of shell it thinks it is in.
# For an “interactive non-login shell”, it reads .bashrc, but for an “interactive login shell” it reads from the first of .bash_profile, .bash_login and .profile (only).
# There is no sane reason why this should be so; it’s just historical. Follows in more detail.

# For Bash, they work as follows. Read down the appropriate column.
# Executes A, then B, then C, etc. The B1, B2, B3 means it executes only the first of those files found.

# +----------------+-----------+-----------+------+
# |                |Interactive|Interactive|Script|
# |                |login      |non-login  |      |
# +----------------+-----------+-----------+------+
# |/etc/profile    |   A       |           |      |
# +----------------+-----------+-----------+------+
# |/etc/bash.bashrc|           |    A      |      |
# +----------------+-----------+-----------+------+
# |~/.bashrc       |           |    B      |      |
# +----------------+-----------+-----------+------+
# |~/.bash_profile |   B1      |           |      |
# +----------------+-----------+-----------+------+
# |~/.bash_login   |   B2      |           |      |
# +----------------+-----------+-----------+------+
# |~/.profile      |   B3      |           |      |
# +----------------+-----------+-----------+------+
# |BASH_ENV        |           |           |  A   |
# +----------------+-----------+-----------+------+
# |                |           |           |      |
# +----------------+-----------+-----------+------+
# |                |           |           |      |
# +----------------+-----------+-----------+------+
# |~/.bash_logout  |    C      |           |      |
# +----------------+-----------+-----------+------+

# you have to add this at the end of the .bash_profile file in home directory if it exists otherwise just copy it to the home directory.

if [ -f ~/.bash_aliases ]; then
    . ~/.bash_aliases
fi

if [ -f ~/.bash_custom ]; then
    . ~/.bash_custom
fi

# Add the following lines to your ~/.bash_profile:
if [[ "$OSTYPE" =~ "darwin" ]] && [[ -f "`which brew`" ]]; then
  if [ -f $(brew --prefix)/etc/bash_completion.d/git-completion.bash ]; then
      . $(brew --prefix)/etc/bash_completion.d/git-completion.bash
  fi
fi

# bash history
export HISTCONTROL=ignoredups

[[ -s "$HOME/.rvm/scripts/rvm" ]] && source "$HOME/.rvm/scripts/rvm" # Load RVM into a shell session *as a function*
