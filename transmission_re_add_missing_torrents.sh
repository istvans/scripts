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


#=============================================================================#
################                    FUNCTIONS                  ################
#=============================================================================#

function show_usage
{
    me=$(basename $0)
    echo """Usage: $me (Options)

 Try to remove and re-add torrents with '$THE_ERROR'.

 -u|--user          USER            a user with access to the transmission deamon
 -p|--password      PASSWORD        the user's passwordto access the transmission deamon
 -t|--torrent       TORRENT_NAME    operate on a single torrent; otherwise try to re-add every missing torrent
 -h|--help                          print this help message

"""
}

function re_add_torrent_from_info
{
    name="$1"; shift
    id="$1"; shift
    info="$1"; shift
    tolerate_error="$1"; shift

    if [[ "$info" =~ $THE_ERROR ]]; then
        torrent_file=$(find . -name "$name*")
        echo "torrent file '$torrent_file' (<- empty == error!)"
        if [[ -z "$torrent_file" ]]; then
            if [[ $tolerate_error = true ]]; then
                echo "WARNING Failed to find the torrent file. SKIPPED"
            else
                exit 1
            fi
        else
            while IFS= read -r line;  do
                if [[ $line =~ Location:[[:space:]]+([^[:space:]].+) ]]; then
                    location="${BASH_REMATCH[1]}"
                    break
                fi
            done <<< $info
            echo "location: '$location' (<- empty == error!)"
            if [[ -z "$location" ]]; then
                if [[ $tolerate_error = true ]]; then
                    echo "WARNING Failed to find the torrent location. SKIPPED"
                else
                    exit 1
                fi
            else
                tmp_torrent_file="$TMP_DIR/$(basename "$torrent_file")"
                echo "back up $torrent_file to $tmp_torrent_file..."
                cp -f "$torrent_file" "$tmp_torrent_file"

                re_add_cmd="$t --add \"$tmp_torrent_file\" --download-dir \"$location\""
                echo "TRY RE_ADD: '$re_add_cmd'"
                eval "$re_add_cmd"

                remove_cmd="$t -t $id --remove"
                echo "REMOVE: '$remove_cmd'"
                eval "$remove_cmd"

                echo "RE_ADD: '$re_add_cmd'"
                eval "$re_add_cmd"
            fi
        fi
    else
        echo SKIPPED
    fi
}

function re_add_torrent_from_line
{
    torrent_line="$1"
    echo "$torrent_line"

    if [[ $torrent_line =~ ^[[:space:]]*([[:digit:]]+).*Stopped[[:space:]]+(.+) ]]; then
        id="${BASH_REMATCH[1]}"
        name="${BASH_REMATCH[2]}"
    else
        echo "Unexpected torrent line format!" >&2 && exit 1
    fi
    echo "id:'$id' name:'$name'"

    info=$($t -t $id --info)
    tolerate_error=true

    re_add_torrent_from_info "$name" "$id" "$info" "$tolerate_error"
}


#=============================================================================#
################                     ARGPARSE                  ################
#=============================================================================#

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
    -h|--help)
        show_usage
        exit 0
        ;;
    *)
        echo "Unknown option $1"
        show_usage
        exit 1
        ;;
  esac
done

[[ -z "$user" ]] && exit 1
[[ -z "$password" ]] && exit 1


#=============================================================================#
################                      MAIN                     ################
#=============================================================================#

t="$tr_cmd $host:$port --auth $user:$password"
cd $torrents_dir

if [[ -n "$torrent" ]]; then
    torrent_file=$(find . -name "$torrent*")
    hash_line=$(transmission-show $torrent_file | fgrep "Hash:")
    torrent_hash=$(echo $hash_line | cut -d' ' -f 2)
    info=$($t -t $torrent_hash --info)
    tolerate_error=false

    re_add_torrent_from_info "$torrent" "$torrent_hash" "$info" "$tolerate_error"
else
    while IFS= read -r line;  do
        re_add_torrent_from_line "$line"
    done < <($t --list | grep Stopped | grep Done)
fi

