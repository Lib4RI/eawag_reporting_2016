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

if [ $# -ne 1 ] || [ ! -f "$1" ]
then
  exit 1
fi

FILE="$1"

EXT="`echo "$FILE" | sed -n '/\./ {s/^.*\.\([^\.]*\)$/\1/; p;}'`"

if [ "$EXT" != "docx" ]
then
  exit 1
fi

WDIR="`pwd`"

TMPDIR="`mktemp -d`"
TMPFILE="`mktemp`"
ZIPTMPFILE="tmp.docx"

unzip -q -d "$TMPDIR" "$FILE"

if [ $? -ne 0 ]
then
  rm "$TMPFILE"
  rm -r "$TMPDIR"
  exit $?
fi

cd "$TMPDIR"

sed -i 's/<\/w:styles>/<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"\/><w:basedOn w:val="DefaultParagraphFont"\/><w:rsid w:val="00F459D1"\/><w:rPr><w:color w:val="0000FF" w:themeColor="hyperlink"\/><w:u w:val="single"\/><\/w:rPr><\/w:style><\/w:styles>/' "$TMPDIR/word/styles.xml"
if [ $? -ne 0 ]
then
  rm "$TMPFILE"
  rm -r "$TMPDIR"
  exit $?
fi

zip -q -r -o "$ZIPTMPFILE" *
if [ $? -ne 0 ]
then
  rm "$TMPFILE"
  rm -r "$TMPDIR"
  exit $?
fi

mv "$ZIPTMPFILE" "$TMPFILE"
if [ $? -ne 0 ]
then
  rm "$TMPFILE"
  rm -r "$TMPDIR"
  exit $?
fi

cd "$WDIR"

cp -p "$TMPFILE" "$FILE"
if [ $? -ne 0 ]
then
  rm "$TMPFILE"
  rm -r "$TMPDIR"
  exit $?
fi

rm "$TMPFILE"
rm -r "$TMPDIR"

exit 0