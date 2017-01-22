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

CAT='/bin/cat'
SED='/bin/sed'
D2U='/usr/bin/dos2unix'
NNL='./nuke_newlines.sed'

THIS="`echo "$0" | $SED '/\// s/.*\/\([^\/]\+\)$/\1/'`"

USAGE="$THIS [-h]"

usage() {
  echo "$USAGE" >&2
  exit 1
}

HELP="NAME
\t$THIS - Preprocess RefWorks \"Tab Delimited with RefID*\"
\t$(echo "$THIS" | $SED 's/./ /g;')   bibliographies

SYNOPSIS
\t   $USAGE

DESCRIPTION
\tParse the standard input, converting from MAC-format to UNIX-format,
\tnuke empty lines, and substitute newlines inside fields with spaces.
\tThe input is assumed to be a tab-separated-values RefWorks-export
\twith the first column containing the RefID. The output is of the
\tsame type, but more usable for post-processing.

OPTIONS
\t-h\tDisplay this message

EXAMPLES
\t   cat export.txt | $THIS > export_preprocessed.txt"

while getopts "h" O
do
  case "$O" in
    h) # show help and exit
      echo "$HELP"
      exit 0
      ;;
    \?) # unknown option: show usage and exit
      usage
      ;;
  esac
done

shift $(($OPTIND - 1))

if [ $# -ne 0 ] # there is still something on the command line
then
  echo "Invalid command line $@"
  usage
fi

$CAT | $D2U -c mac | $NNL