#! /bin/bash -l

function trace_start() {
    echo
    echo "************************ start $1 ************************"
    echo
}

function trace() {
  echo "----------------------------------------------------------------------"
  echo `date +%Y-%m-%dT%H:%M:%S` " - " $1
  echo ""
}

function trace_end() {
    echo "************************ end $1 ************************"
}