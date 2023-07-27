#!/bin/bash
#exec >/dev/tty 2>/dev/tty </dev/tty && screen -S display "bash"
#
#
script /dev/null
screen -wipe
sleep 1
screen -d -m -S display "bash"
rm -f out.txt || :
sleep 1
screen -S display -p 0 -X stuff "source wine_display.sh"
sleep 4
screen -S display -p 0 -X stuff "./do_vbs.sh > out.tmp.txt && mv out.tmp.txt out.txt"

exec sleep infinity

