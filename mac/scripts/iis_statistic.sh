#! /bin/bash
# usage example: iis_statisti.sh "hotellist*"
# description
# calculates which url in iis log took more than e.g. 10s but less than 20s and etc.
# output format: 3393 < 10 > 383 < 20 > 84 < 40 > 3 < 60 > 3 - 3866 - 20130613


files=(`find ./ -type f -name "$1" | sort`)

for file in ${files[*]}; do
  #echo file name: $file
  result_file=`basename $file`
  awk 'BEGIN {count=0;a=0;a10=0;a20=0;a40=0;a60=0}
  { 
    count=count+1

    if (int($22) > 60000) {
      a60=a60+1
    } else if (int($22) > 40000) {
      a40=a40+1
    } else if (int($22) > 20000) {
      a20=a20+1
    } else if (int($22) > 10000) {
      a10=a10+1
    } else {
      a=a+1
    }
  }
  END {
    gsub(/[^[0-9]/,"", FILENAME)
    print a " < 10s > " a10 " < 20s > " a20 " < 40s > " a40 " < 60s > " a60 " - " count " - " FILENAME 
  }' $file
done
