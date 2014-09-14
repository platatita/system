#! /bin/bash

# it replaces all these characters [ -()+\[\]] in the folder or file name with underscore
# options:
# d - replace folders names [default]
# f - replace files names

# usage
# * replace all subfolders of ~/tmp folder: ./translate_path ~/tmp 
# * replace all subfolders of current dir: ./translate_path .
# * replace all file names of current dir: ./translate_path . f

# example of output:
# ./Percy Jackson and the Olympians -> ./percy_jackson_and_the_olympians
# ./simple folder -> ./simple_folder

if [ $# -lt 1 ]; then
  echo "usage"
  echo "./translate_path folder [folder (default) | in files (f)]"
  echo "examples:"
  echo "* replace all subfolders of ~/tmp folder: ./translate_path ~/tmp" 
  echo "* replace all subfolders of current dir:  ./translate_path ."
  echo "* replace all file names of current dir:  ./translate_path . f"
  exit 0
fi

FIND_DIR="$1"
FIND_TYPE=$2

if [ -z $FIND_TYPE ]; then
  FIND_TYPE="d"
fi

# find $FIND_DIR -type $FIND_TYPE -iname "* *" -not -path "." | awk '{org = $0; gsub(/ /,"_", $0); printf "\"%s\" %s\n", org, tolower($0)}' | xargs -n2 mv -v

find $FIND_DIR -type $FIND_TYPE -not -path "." | awk -v FIND_TYPE=$FIND_TYPE '
  {
    org_0 = $0; 
    gsub(/[ -()+\[\]]/,"_", $0);

    split_0_count = split($0, split_0, ".");
    if (split_0_count > 1)
    {
      file = FIND_TYPE == "d" ? file = 0 : file = 1;

      result = split_0[1];

      for (i = 2; i <= split_0_count - file; i++)
      {
        result = result "_" tolower(split_0[i]);
      }

      if (file == 1)
      {
        result = result "." split_0[split_0_count];
      }
    }
    else
    {
      result = $0;
    }
    
    if (result != org_0)
    {
      printf "\"%s\" %s\n", org_0, result;
    }

  }' | xargs -n2 mv -v
