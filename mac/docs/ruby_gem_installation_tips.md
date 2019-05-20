# to install ruby 1.9.3-p551 on mac os
rvm install 1.9.3 --with-gcc=clang

# to install therubyracer when some errors occurres
# therubyracer install command
gem install therubyracer
# or with specifying version
gem install therubyracer -v '0.12.1' -- --with-system-v8


# then
# libv8 install command
gem install libv8 -v '3.16.14.3' -- --with-system-v8


# to install capybara-webkit-1.3.0
brew install qt5
gem install capybara-webkit -v '1.3.0' --source 'https://rubygems.org/'


# to install capybara-webkit-1.6.0 e.g. for blenderweb
brew install qt5
brew link --force qt5
gem install capybara-webkit -v '1.6.0'


# to install nokogiri ruby 1.9.3
brew install libiconv
brew link libiconv
gem install nokogiri -v '1.6.3.1' --source 'https://rubygems.org/' -- --with-iconv-dir=/usr/local/Cellar/libiconv/1.16


