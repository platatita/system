# to install therubyracer when some errors occurres

# therubyracer install command
gem install therubyracer
# or with specifying version
gem install therubyracer -v '0.12.1' -- --with-system-v8


# then
# libv8 install command
gem install libv8 -v '3.16.14.3' -- --with-system-v8


# to install capybara-webkit-1.6.0 e.g. for blenderweb
brew install qt@5.5
brew link --force qt55
gem install capybara-webkit -v '1.6.0'
