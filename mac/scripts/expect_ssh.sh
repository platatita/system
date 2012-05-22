#! /usr/bin/expect -f
# the script presents you how to use the 'expect' to autologin into remote server and provide password.

set timeout 60

set user  [lindex $argv 0]
set serv  [lindex $argv 1]
set pass  [lindex $argv 2]
set comm  [lindex $argv 3]

spawn ssh $user@$serv
expect "password:" { send "$pass\r" }
send "$comm\r"

# switch off below two lines if you want use interact mode
send "exit\r"
expect eof

# switch on if you want to put any command after login
# interact