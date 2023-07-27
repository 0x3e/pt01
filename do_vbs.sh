#!/usr/bin/env bash
#

wine_exe='wine'

cat vbs/errors.vbs > all.vbs
find vbs/classes/ -iname '*.vbs' -exec cat {} \; >> all.vbs
cat vbs/load_globals.vbs >> all.vbs
$wine_exe cscript.exe all.vbs | tee out/tmp_run.txt
echo "end run"
echo
echo
echo
echo "diff"
diff -c4 out/expected_run.txt out/tmp_run.txt
