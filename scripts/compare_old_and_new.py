#!/usr/bin/python
# coding=utf-8

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

### load stuff we need
import sys
### the following two lines are a dirty hack and are discouraged! we save time by not looking for another solution. see also http://stackoverflow.com/questions/2276200/changing-default-encoding-of-python and https://anonbadger.wordpress.com/2015/06/16/why-sys-setdefaultencoding-will-break-code/
reload(sys)
sys.setdefaultencoding("utf8")
from optparse import OptionParser
import os
import codecs
import re
import csv
from collections import OrderedDict
from cStringIO import StringIO

### some constants

STDIN   = sys.stdin  #'/dev/stdin'
STDOUT  = sys.stdout #'/dev/stdout'

testing = True
testing = False ### uncomment this line if you are done testing

### re-direct all output to stderr

oldstdout = sys.stdout
sys.stdout = sys.stderr

### parse the command line

usage = "Usage: %prog [-v] -l lastyearfile -o oldfile -n newfile"
parser = OptionParser(usage)
parser.add_option("-v", "--verbose",
  action = "store_true", dest = "verbose", default = False,
  help = "Show what I'm doing [default=false]")
parser.add_option("-l", "--lastyearfile", nargs = 1,
  action = "store", dest = "lyfile", metavar = "lyfile",
  help = "Last (2015) year's tab delimitted file (without RefID)")
parser.add_option("-o", "--oldfile", nargs = 1,
  action = "store", dest = "oldfile", metavar = "oldfile",
  help = "The old (2015) tab delimitted file (with RefID)")
parser.add_option("-n", "--newfile", nargs = 1,
  action = "store", dest = "newfile", metavar = "newfile",
  help = "The new (2016) tab delimitted file (with RefID)")
(opts, args) = parser.parse_args()

verbose = opts.verbose

if not opts.lyfile:
  print "Error: You need to specify the last year (2015) file!"
  exit(1)
else:
  lyfile = opts.lyfile

if not opts.oldfile:
  print "Error: You need to specify the old (2015) file!"
  exit(1)
else:
  oldfile = opts.oldfile

if not opts.newfile:
  print "Error: You need to specify the new (2016) file!"
  exit(1)
else:
  newfile = opts.newfile

if not os.access(lyfile, os.R_OK):
  parser.error("Cannot read from " + lyfile + "!")

if not os.access(oldfile, os.R_OK):
  parser.error("Cannot read from " + oldfile + "!")

if not os.access(newfile, os.R_OK):
  parser.error("Cannot read from " + newfile + "!")

# following two definitions from https://docs.python.org/2/library/csv.html#examples
def unicode_csv_reader(unicode_csv_data, dialect=csv.excel, **kwargs):
    # csv.py doesn't do Unicode; encode temporarily as UTF-8:
    csv_reader = csv.reader(utf_8_encoder(unicode_csv_data),
                            dialect=dialect, **kwargs)
    for row in csv_reader:
        # decode UTF-8 back to Unicode, cell by cell:
        yield [unicode(cell, 'utf-8') for cell in row]
def utf_8_encoder(unicode_csv_data):
    for line in unicode_csv_data:
        yield line.encode('utf-8')

lyheader = None
lyvals = []
with codecs.open(lyfile, "rb", "utf8") as lyfd:
  csvreader = unicode_csv_reader(lyfd, delimiter = "\t")
  for row in csvreader:
    if not lyheader:
      lyheader = row
      continue
    result = OrderedDict()
    for i in range(len(row)):
      result.update({lyheader[i] : row[i]})
    lyvals.append(result)
  lyfd.close()

oldheader = None
oldvals = []
with codecs.open(oldfile, "rb", "utf8") as oldfd:
  csvreader = unicode_csv_reader(oldfd, delimiter = "\t")
  for row in csvreader:
    if not oldheader:
      oldheader = row
      continue
    result = OrderedDict()
    for i in range(len(row)):
      result.update({oldheader[i] : row[i]})
    oldvals.append(result)
  oldfd.close()

newheader = None
newvals = []
with codecs.open(newfile, "rb", "utf8") as newfd:
  csvreader = unicode_csv_reader(newfd, delimiter = "\t")
  for row in csvreader:
    if not newheader:
      newheader = row
      continue
    result = OrderedDict()
    for i in range(len(row)):
      result.update({newheader[i] : row[i]})
    newvals.append(result)
  newfd.close()

#

def myequal(a, b):
  return (a and a not in ["", " "] and b and b not in ["", " "] and a == b)

oldfoundinly  = [] # indices refer to lyfile
oldnotfoundinly = [] # contains refids
lynotfoundinold = [] # indices refer to lyfile
newfoundinly = [] # indices refer to lyfile
newfoundinold = [] # contains refids

if verbose:
  print "Comparing " + lyfile + " and " + oldfile + "."

firstline = True
for oval in oldvals:
  ono = None
  onop = None
  ourl = oval["URL"]
  onotes = oval['Notes']
  if ourl:
    try:
      onop = int(re.sub("http://library.eawag.ch/eawag-publications/(pdf|openaccess)/Eawag_([0-9]*).pdf", r'\2', ourl, flags = re.I))
    except ValueError:
      pass
  try:
    ono = int(onotes)
  except ValueError:
    if onop:
      ono = onop
    pass
  if ono and onop and ono != onop:
    print "Warning: Mismatch in RefID." + oval["ID"] + "/Eawag-No." + str(ono) + " (the pdf suggests " + str(onop) + ")."
  found = False
  linksalerted = False
  for i in range(len(lyvals)):
    lval = lyvals[i]
    lno = None
    lnop = None
    lurl = lval["URL"]
    lnotes = lval['Notes']
    if lurl:
      try:
        lnop = int(re.sub("http://library.eawag.ch/eawag-publications/(pdf|openaccess)/Eawag_([0-9]*).pdf", r'\2', lurl, flags = re.I))
      except ValueError:
        pass
    try:
      lno = int(lnotes)
    except ValueError:
      if lnop:
        lno = lnop
      pass
    if firstline and lno and lnop and lno != lnop:
      print "Warning: Mismatch in Eawag-No." + str(lno) + " (the pdf suggests " + str(lnop) + ")."
    if lno and ono and lno == ono:
#      if verbose:
#        print "Info: Found RefID." + oval['ID'] + " to be in both lists (Eawag-No.=" + lnotes + ")."
      oldfoundinly.append(i)
      found = True
      break
    elif myequal(lurl, ourl):
      if testing:
        print "\"" + lnotes + "\"!=\"" + onotes + "\" || \"" + lurl + "\"!=\"" + ourl + "\""
      if verbose:
        print "Info: Found RefID." + oval["ID"] + " to be in both lists (URL=" + ourl + ")."
      oldfoundinly.append(i)
      found = True
      break
    elif myequal(lval["Authors, Primary"], oval["Authors, Primary"]) and myequal(lval["Title Primary"], oval["Title Primary"]):
      if testing:
        print "\"" + lnotes + "\"!=\"" + onotes + "\" || \"" + lurl + "\"!=\"" + ourl + "\""
      if verbose:
        print "Info: Found RefID." + oval["ID"] + " to be in both lists (Authors=" + oval["Authors, Primary"] + " and Title=" + oval["Title Primary"] + ")."
      oldfoundinly.append(i)
      found = True
      break
    elif myequal(lval["Links"], oval["Links"]):
      if testing:
        print "\"" + lnotes + "\"!=\"" + onotes + "\" || \"" + lurl + "\"!=\"" + ourl + "\""
      if verbose and not linksalerted:
        print "Notice: RefID." + oval["ID"] + " might be in both lists (Links=" + oval["Links"] + "), but I am not sure... Check manually!"
      linksalerted = True
  firstline = False
  if not found:
    oldnotfoundinly.append(oval["ID"])
#

lynotfoundinold = sorted(set(range(len(lyvals)))-set(oldfoundinly), key=lambda t: int(t))

if verbose:
  print "Comparing " + lyfile + " and " + newfile + "."

firstline = True
for nval in newvals:
  nno = None
  nnop = None
  nurl = nval["URL"]
  nnotes = nval['Notes']
  if nurl:
    try:
      nnop = int(re.sub("http://library.eawag.ch/eawag-publications/(pdf|openaccess)/Eawag_([0-9]*).pdf", r'\2', nurl, flags = re.I))
    except ValueError:
      pass
  try:
    nno = int(nnotes)
  except ValueError:
    if nnop:
      nno = nnop
    pass
  if nno and nnop and nno != nnop:
    print "Warning: Mismatch in RefID." + nval["ID"] + "/Eawag-No." + str(nno) + " (the pdf suggests " + str(nnop) + ")."
  linksalerted = False
  for i in range(len(lyvals)):
    lval = lyvals[i]
    lno = None
    lnop = None
    lurl = lval["URL"]
    lnotes = lval['Notes']
    if lurl:
      try:
        lnop = int(re.sub("http://library.eawag.ch/eawag-publications/(pdf|openaccess)/Eawag_([0-9]*).pdf", r'\2', lurl, flags = re.I))
      except ValueError:
        pass
    try:
      lno = int(lnotes)
    except ValueError:
      if lnop:
        lno = lnop
      pass
#    if firstline and lno and lnop and lno != lnop:
#      print "Warning: Mismatch in Eawag-No." + str(lno) + " (the pdf suggests " + str(lnop) + ")."
    if lno and nno and lno == nno:
      if verbose:
        print "Info: Found RefID." + nval['ID'] + " to be in both lists (Eawag-No.=" + nnotes + ")."
      newfoundinly.append(i)
      break
    elif myequal(lurl, nurl):
      if testing:
        print "\"" + lnotes + "\"!=\"" + nnotes + "\" || \"" + lurl + "\"!=\"" + nurl + "\""
      if verbose:
        print "Info: Found RefID." + nval["ID"] + " to be in both lists (URL=" + nurl + ")."
      newfoundinly.append(i)
      break
    elif myequal(lval["Authors, Primary"], nval["Authors, Primary"]) and myequal(lval["Title Primary"], nval["Title Primary"]):
      if testing:
        print "\"" + lnotes + "\"!=\"" + nnotes + "\" || \"" + lurl + "\"!=\"" + nurl + "\""
      if verbose:
        print "Info: Found RefID." + nval["ID"] + " to be in both lists (Authors=" + nval["Authors, Primary"] + " and Title=" + nval["Title Primary"] + ")."
      newfoundinly.append(i)
      break
    elif myequal(lval["Links"], nval["Links"]):
      if testing:
        print "\"" + lnotes + "\"!=\"" + nnotes + "\" || \"" + lurl + "\"!=\"" + nurl + "\""
      if verbose and not linksalerted:
        print "Notice: RefID." + nval["ID"] + " might be in both lists (Links=" + nval["Links"] + "), but I am not sure... Check manually!"
      linksalerted = True
  firstline = False
#

if verbose:
  print "Comparing " + oldfile + " and " + newfile + "."

firstline = True
for nval in newvals:
  nno = None
  nnop = None
  nurl = nval["URL"]
  nnotes = nval['Notes']
  if nurl:
    try:
      nnop = int(re.sub("http://library.eawag.ch/eawag-publications/(pdf|openaccess)/Eawag_([0-9]*).pdf", r'\2', nurl, flags = re.I))
    except ValueError:
      pass
  try:
    nno = int(nnotes)
  except ValueError:
    if nnop:
      nno = nnop
    pass
#  if nno and nnop and nno != nnop:
#    print "Warning: Mismatch in RefID." + nval["ID"] + "/Eawag-No." + str(nno) + " (the pdf suggests " + str(nnop) + ")."
  linksalerted = False
  for oval in oldvals:
    ono = None
    onop = None
    ourl = oval["URL"]
    onotes = oval['Notes']
    if ourl:
      try:
        onop = int(re.sub("http://library.eawag.ch/eawag-publications/(pdf|openaccess)/Eawag_([0-9]*).pdf", r'\2', ourl, flags = re.I))
      except ValueError:
        pass
    try:
      ono = int(onotes)
    except ValueError:
      if onop:
        ono = onop
      pass
#    if firstline and ono and onop and ono != onop:
#      print "Warning: Mismatch in Eawag-No." + str(ono) + " (the pdf suggests " + str(onop) + ")."
    if ono and nno and ono == nno:
      if verbose:
        print "Info: Found RefID." + nval['ID'] + "/" + oval['ID'] + " to be in both lists (Eawag-No.=" + nnotes + "/" + onotes + ")."
      newfoundinold.append(nval["ID"])
      break
    elif myequal(ourl, nurl):
      if testing:
        print "\"" + onotes + "\"!=\"" + nnotes + "\" || \"" + ourl + "\"!=\"" + nurl + "\""
      if verbose:
        print "Info: Found RefID." + nval["ID"] + "/" + oval['ID'] + " to be in both lists (URL=" + nurl + ")."
      newfoundinold.append(nval["ID"])
      break
    elif myequal(oval["Authors, Primary"], nval["Authors, Primary"]) and myequal(oval["Title Primary"], nval["Title Primary"]):
      if testing:
        print "\"" + onotes + "\"!=\"" + nnotes + "\" || \"" + ourl + "\"!=\"" + nurl + "\""
      if verbose:
        print "Info: Found RefID." + nval["ID"] + "/" + oval['ID'] + " to be in both lists (Authors=" + nval["Authors, Primary"] + " and Title=" + nval["Title Primary"] + ")."
      newfoundinold.append(nval["ID"])
      break
    elif myequal(oval["Links"], nval["Links"]):
      if testing:
        print "\"" + onotes + "\"!=\"" + nnotes + "\" || \"" + ourl + "\"!=\"" + nurl + "\""
      if verbose and not linksalerted:
        print "Notice: RefID." + nval["ID"] + "/" + oval['ID'] + " might be in both lists (Links=" + nval["Links"] + "), but I am not sure... Check manually!"
      linksalerted = True
  firstline = False
#

if verbose:
  print "2015 entries found solely in " + lyfile + " (" + str(len(lynotfoundinold)) + "):"
  print "  ", [str(k) + ": " + lyvals[k]["Notes"] for k in lynotfoundinold]
  print "2015 entries found solely in " + oldfile + " (" + str(len(oldnotfoundinly)) + "):"
  print "  ", sorted(oldnotfoundinly, key=lambda t: int(t))
  print "2016 entries found in " + lyfile + " (" + str(len(newfoundinly)) + "):"
  print "  ", [str(k) + ": " + lyvals[k]["Notes"] for k in newfoundinly]
  print "2016 entries found in " + oldfile + " (" + str(len(newfoundinold)) + "):"
  print "  ", sorted(newfoundinold, key=lambda t: int(t))


