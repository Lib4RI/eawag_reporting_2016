#!/bin/sh

###
 # Copyright (c) 2016, 2017 d-r-p (Lib4RI) <d-r-p@users.noreply.github.com>
 #
 # Permission to use, copy, modify, and distribute this software for any
 # purpose with or without fee is hereby granted, provided that the above
 # copyright notice and this permission notice appear in all copies.
 #
 # THE SOFTWARE IS PROVIDED "AS IS" AND THE AUTHOR DISCLAIMS ALL WARRANTIES
 # WITH REGARD TO THIS SOFTWARE INCLUDING ALL IMPLIED WARRANTIES OF
 # MERCHANTABILITY AND FITNESS. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR
 # ANY SPECIAL, DIRECT, INDIRECT, OR CONSEQUENTIAL DAMAGES OR ANY DAMAGES
 # WHATSOEVER RESULTING FROM LOSS OF USE, DATA OR PROFITS, WHETHER IN AN
 # ACTION OF CONTRACT, NEGLIGENCE OR OTHER TORTIOUS ACTION, ARISING OUT OF
 # OR IN CONNECTION WITH THE USE OR PERFORMANCE OF THIS SOFTWARE.
###

REPDIR="`cat $0 | sed -n '21, /prepare_reports/ {s/.*prepare_reports.py.*-o \(.\+\) -y.*/\1/ p}'`"

if [ ! -d "$REPDIR" ]
then
  mkdir -p "$REPDIR" || exit $?
fi

cat 20170103_JIF_Full_WoS_Export.csv | sed '1 d; $ d;' | sed "`cat 20170103_JIF_Full_WoS_Export.csv | sed '1 d; $ d;' | sed -n '/^\([^,]*,\)\{3\},/! d; =;'` s/^\(\([^,]*,\)\{3\}\)/\10263-0923/" > 20170103_JIF_Full_WoS_Export.consolidated.csv
if [ $? -ne 0 ]
then
  exit $?
fi

iconv -f windows-1252 -t utf-8 Reporting_20160105_AllReferences_all.txt | dos2unix > Reporting_20160105_AllReferences_all.utf8.txt
if [ $? -ne 0 ]
then
  exit $?
fi

./preprocess_exports.sh < 2015_tab.txt > 2015_tab.preprocessed.txt 
if [ $? -ne 0 ]
then
  exit $?
fi
./preprocess_exports.sh < 2016_tab.txt > 2016_tab.preprocessed.txt 
if [ $? -ne 0 ]
then
  exit $?
fi

./preprocess_exports.sh < 2015_reports_tab.txt > 2015_reports_tab.preprocessed.txt 
if [ $? -ne 0 ]
then
  exit $?
fi
./preprocess_exports.sh < 2016_reports_tab.txt > 2016_reports_tab.preprocessed.txt 
if [ $? -ne 0 ]
then
  exit $?
fi

cat 2015_tab.preprocessed.txt > 2015_tab.all.txt && tail -n +2 2015_reports_tab.preprocessed.txt >> 2015_tab.all.txt
if [ $? -ne 0 ]
then
  exit $?
fi
cat 2016_tab.preprocessed.txt > 2016_tab.all.txt && tail -n +2 2016_reports_tab.preprocessed.txt >> 2016_tab.all.txt
if [ $? -ne 0 ]
then
  exit $?
fi

PYTHONIOENCODING=utf8 python compare_old_and_new.py -v -l Reporting_20160105_AllReferences_all.utf8.txt -o 2015_tab.all.txt -n 2016_tab.all.txt 2> compare_old_and_new.log 
if [ $? -ne 0 ]
then
  exit $?
fi

PYTHONIOENCODING=utf8 python dora_lookup.py -v < 2016_tab.all.txt > 2016_tab.all.withdorapids.txt 2> dora_lookup.log
if [ $? -ne 0 ]
then
  exit $?
fi

PYTHONIOENCODING=utf8 python prepare_reports.py -v -p 2016_report -t "Report 2016" -o 2016_reports -y 2015 -j 20170103_JIF_Full_WoS_Export.consolidated.csv < 2016_tab.all.withdorapids.txt 2> prepare_reports.log
if [ $? -ne 0 ]
then
  exit $?
fi

for f in "$REPDIR"/*.docx
do
  ./amend_docx_styles.sh "$f"
  if [ $? -ne 0 ]
  then
    exit $?
  fi
done

exit 0