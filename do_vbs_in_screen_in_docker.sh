#!/usr/bin/env bash
set -e

rm -f out.txt || :
docker compose exec wine-ubuntu sh -c 'screen -S display -p 0 -X stuff "./do_vbs.sh > out.tmp.txt && mv out.tmp.txt out.txt"'

i=0
while [ ! -f out.txt ]
do
  ((i+=1))
  sleep 0.1
  if ((i>300)); then 
    echo "badness" 
    exit 11
  fi
done

cat out.txt

