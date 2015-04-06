#! /bin/bash

# below command dispaly result of lsof -FT and repeat this command +r till you interup it
# lsof -FT +r | awk 'BEGIN{FS="="};{if ($0 ~ /^p/) {pid=$0}; if ($0 ~ /^T/) { if (int($2) == 0) {print pid "->" $1 "=" $2} }}'


# below block of code does the same like code above but adds line separator after each iteration.
declare -i counter=0

while true; do
  # below command assign the result of command lsof -FT to variable
  lsof_result=`lsof -FT` 
  
  echo $lsof_result | awk 'BEGIN{FS="=";RS=" "};{if ($0 ~ /^p/) {pid=$0}; if ($0 ~ /^T/) { if (int($2) > 0) {print pid "->" $1 "=" $2} }}'
  
  echo `date "+%Y-%m-%d %H:%M:%S"` - $counter "-------------------------------"
  counter=counter+1
  sleep 5
done
