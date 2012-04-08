#! /bin/bash -l

# goes through each subfolder to find the '.git' folder and next do 'git pull' command

function get_git_directories {
  ls -d */.git
}

function iterate_through_git_repo {
  for (( i=0; i<$DIRS_LEN;i++ )) do
    local repo=${DIRS[${i}]}
    local proj_dir=${repo/\/\.git*/}
    echo ---------- proj_dir: $proj_dir, git_dir: $repo ---------- 
    do_git_pull $proj_dir
  done
}

function do_git_pull {
    cd $1
    echo `pwd`
    sleep 1
    git pull
    cd $WORKDIR
}

WORKDIR=`pwd`
DIRS=(`get_git_directories`)
DIRS_LEN=${#DIRS[@]}

if [ $DIRS_LEN -gt 0 ]; then
  echo found $DIRS_LEN git repos
  iterate_through_git_repo
else
  echo ---------- did not find any git repos ----------
fi

