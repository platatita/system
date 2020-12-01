#! /bin/bash

# copy all files starting with '.' dot to the HOME dir
find . -type f -name ".*" -exec cp -v {} ~/ \;

