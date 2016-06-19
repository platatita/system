#! /bin/bash

if [ -f ~/.bash_aliases_anixe ]; then
    source ~/.bash_aliases_anixe
fi

# github repos aliases
alias system='cd ~/sources/git/system'
alias dev='cd ~/sources/git/dev'
alias skydrive='cd ~/SkyDrive'
alias dropbox='cd ~/Dropbox'
alias download='cd ~/Download'

# media
alias media='cd /var/data/media'

# launching sublime alias
alias sublime='~/.sublime_run.sh'

# mongo
alias pmongo='cd /opt/mongodb/current'

# redis
alias predis='cd /opt/redis/current'

# passenger
alias ppassenger=' cd /opt/passenger/current'

# memcached
alias pmemcached='cd /opt/memcached/current'

# git
alias gs='git status'
alias gp='git pull'
alias gb='git branch'
alias gl='git log --color'

# bundler
alias be='bundle exec'
alias ber='bundle exec rake -T'
alias bers='bundle exec rake spec'
alias bert='bundle exec rake test'
alias berc='bundle exec rake cukes'
alias berf='bundle exec rake features'

# system
alias ll='ls -asl'
#alias dtree="find . -type d -print | sed -e 's;[^/]*/;|____;g;s;____|; |;g'"
alias dtree=fdtree
alias tree="find . -print | sed -e 's;[^/]*/;|____;g;s;____|; |;g'"
alias hosts='cat /etc/hosts'
alias vhosts='sudo vim /etc/hosts'

function fdtree {
  depth=$1
  if [ -z $depth ]; then
    depth=3
  fi
  find . -maxdepth $depth -type d -print | sed -e 's;[^/]*/;|____;g;s;____|; |;g'
}

# networksetup
alias ns='networksetup'
alias nsc='networksetup -printcommands'
alias list-pref-wifi='ns -listpreferredwirelessnetworks en0'

# airport
alias wifi-scan='airport -s'

OSNAME=`uname`
if [ $OSNAME = 'Linux' ]; then
  alias vbash='vim ~/.bashrc'
  alias sbash='source ~/.bashrc'
else
  alias vbash='vim ~/.bash_profile'
  alias sbash='source ~/.bash_profile'
fi
