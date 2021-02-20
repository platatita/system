**Mac OS X**

**Info about the '/var/folders'** in Mac Os: http://www.magnusviri.com/what-is-var-folders.html

```bash
# Get the path to the 'cache' directory where complied Arduino libraries are stored
getconf DARWIN_USER_CACHE_DIR
/var/folders/vz/xdqgcrp57637bbxj5pnjrmqr0000gp/C/

# Get the path to the 'tmp' directory
getconf DARWIN_USER_TEMP_DIR
/var/folders/vz/xdqgcrp57637bbxj5pnjrmqr0000gp/T/

# And find 'Arduino.app'
find /var/folders/vz/xdqgcrp57637bbxj5pnjrmqr0000gp/T -type d -name "Arduino.app"
/var/folders/vz/xdqgcrp57637bbxj5pnjrmqr0000gp/T/AppTranslocation/2E98E07D-A8FF-4BF0-BE8F-34133CDA3EC3/d/Arduino.app

# And then find 'libraries' folder
find /var/folders/vz/xdqgcrp57637bbxj5pnjrmqr0000gp/T/AppTranslocation/2E98E07D-A8FF-4BF0-BE8F-34133CDA3EC3/d/Arduino.app -type d -name "libraries"
/var/folders/vz/xdqgcrp57637bbxj5pnjrmqr0000gp/T/AppTranslocation/2E98E07D-A8FF-4BF0-BE8F-34133CDA3EC3/d/Arduino.app/Contents/Java/libraries

# And then list the content
ls -asl /var/folders/vz/xdqgcrp57637bbxj5pnjrmqr0000gp/T/AppTranslocation/2E98E07D-A8FF-4BF0-BE8F-34133CDA3EC3/d/Arduino.app/Contents/Java/libraries

# Adafruit_Circuit_Playground
# Bridge
# Esplora
# Ethernet
# Firmata
# GSM
# Keyboard
# LiquidCrystal
# Mouse
# RobotIRremote
# Robot_Control
# Robot_Motor
# SD
# Servo
# SpacebrewYun
# Stepper
# TFT
# Temboo
# WiFi
```
