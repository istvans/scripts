#!/usr/bin/env bash
# Helper script to alleviate the pain that this transmission issue causes:
# https://github.com/transmission/transmission/issues/83
# Try to use this script if you see $THE_ERROR. Good luck!

set -euo pipefail  # enable strict mode

THE_ERROR='Error: No data found! Ensure your drives are connected or use "Set Location". To re-download, remove the torrent and re-add it.'
TMP_DIR=/tmp

tr_cmd=transmission-remote
host=localhost
port=9091
user=
password=
torrents_dir=$HOME/.config/transmission-daemon/torrents
torrent=

while [[ $# -gt 0 ]]; do
  case $1 in
    -u|--user)
      user="$2"
      shift # past argument
      shift # past value
      ;;
    -p|--password)
      password="$2"
      shift # past argument
      shift # past value
      ;;
    -t|--torrent)
      torrent="$2"
      shift # past argument
      shift # past value
      ;;
    *)
      echo "Unknown option $1"
      exit 1
      ;;
  esac
done

[[ -z "$user" ]] && exit 1
[[ -z "$password" ]] && exit 1

t="$tr_cmd $host:$port --auth $user:$password"

cd $torrents_dir

if [[ -n "$torrent" ]]; then
    echo "Filtering for a single torrent: '$torrent'"
    torrent_filter="| grep -P '\s$torrent$'"
else
    torrent_filter=
fi

function re_add_torrent
{
    line="$1"
    echo "$line"
    IFS=', ' read -r -a fields <<< "$line"
    echo "'${fields[0]}' '${fields[-1]}'"
    id=$(echo ${fields[0]} | sed 's#*##g')
    name="${fields[-1]}"
    echo "id:'$id' name:'$name'"

    info=$($t -t $id --info)
    if [[ "$info" =~ $THE_ERROR ]]; then
        torrent_file=$(ls $name*)
        echo "torrent file '$torrent_file'"

        while IFS= read -r line;  do
            if [[ $line =~ Location:[[:space:]]+([^[:space:]].+) ]]; then
                location="${BASH_REMATCH[1]}"
                break
            fi
        done <<< $info
        echo "location: '$location' (<- empty == error!)"
        [[ -z "$location" ]] && exit 1

        tmp_torrent_file="$TMP_DIR/$(basename $torrent_file)"
        echo "backup $torrent_file to $tmp_torrent_file..."
        cp -f $torrent_file $tmp_torrent_file

        re_add_cmd="$t --add $tmp_torrent_file --download-dir $location"
        echo "TRY RE_ADD: '$re_add_cmd'"
        eval "$re_add_cmd"

        remove_cmd="$t -t $id --remove"
        echo "REMOVE: '$remove_cmd'"
        eval "$remove_cmd"

        echo "RE_ADD: '$re_add_cmd'"
        eval "$re_add_cmd"
    else
        echo SKIPPED
    fi

}

while IFS= read -r line;  do
    re_add_torrent "$line"
done < <(eval "$t --list | grep Stopped | grep Done${torrent_filter}")

