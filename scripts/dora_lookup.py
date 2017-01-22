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
import requests
import re
import csv
from collections import OrderedDict
from cStringIO import StringIO

### some constants

STDIN   = sys.stdin  #'/dev/stdin'
STDOUT  = sys.stdout #'/dev/stdout'

### re-direct all output to stderr

oldstdout = sys.stdout
sys.stdout = sys.stderr

### parse the command line

usage = "Usage: %prog [-v] [-o outfile] [-i infile | infile]"
parser = OptionParser(usage)
parser.add_option("-v", "--verbose",
  action = "store_true", dest = "verbose", default = False,
  help = "Show what I'm doing [default=false]")
parser.add_option("-i", "--infile", nargs = 1,
  action = "store", dest = "infile", metavar = "infile",
  help = "The input file (tab delimited with RefID)")
parser.add_option("-o", "--outfile", nargs = 1,
  action = "store", dest = "outfile", metavar = "outfile",
  help = "The output file (tab delimited with RefID and DORA PID)")

(opts, args) = parser.parse_args()

verbose = opts.verbose

if verbose:
  print "Parsing command line..."

fromstdin = True
tostdout = True

if len(args) > 1 or (len(args) == 1 and opts.infile):
  parser.error("Expected only one input file!")

if args:
  infile = args[0]
  fromstdin = False
elif opts.infile:
  infile = opts.infile
  fromstdin = False
else:
  infile = STDIN

if opts.outfile:
  outfile = opts.outfile
  tostdout = False
else:
  outfile = STDOUT

if not fromstdin and not os.access(infile, os.R_OK):
  parser.error("Cannot read from " + infile + "!")

if not tostdout and not os.access(outfile, os.W_OK):
  parser.error("Cannot write to " + outfile + "!")

# following three definitions from https://docs.python.org/2/library/csv.html#examples
class UTF8Recoder:
    """
    Iterator that reads an encoded stream and reencodes the input to UTF-8
    """
    def __init__(self, f, encoding):
        self.reader = codecs.getreader(encoding)(f)
    def __iter__(self):
        return self
    def next(self):
        return self.reader.next().encode("utf-8")
class UnicodeReader:
    """
    A CSV reader which will iterate over lines in the CSV file "f",
    which is encoded in the given encoding.
    """
    def __init__(self, f, dialect=csv.excel, encoding="utf-8", **kwds):
        f = UTF8Recoder(f, encoding)
        self.reader = csv.reader(f, dialect=dialect, **kwds)
    def next(self):
        row = self.reader.next()
        return [unicode(s, "utf-8") for s in row]
    def __iter__(self):
        return self
class UnicodeWriter:
    """
    A CSV writer which will write rows to CSV file "f",
    which is encoded in the given encoding.
    """
    def __init__(self, f, dialect=csv.excel, encoding="utf-8", **kwds):
        # Redirect output to a queue
        self.queue = StringIO()
        self.writer = csv.writer(self.queue, dialect=dialect, **kwds)
        self.stream = f
        self.encoder = codecs.getincrementalencoder(encoding)()
    def writerow(self, row):
        self.writer.writerow([(s if s else "").encode("utf-8") for s in row])
        # Fetch UTF-8 output from the queue ...
        data = self.queue.getvalue()
        data = data.decode("utf-8")
        # ... and reencode it into the target encoding
        data = self.encoder.encode(data)
        # write to the target stream
        self.stream.write(data)
        # empty queue
        self.queue.truncate(0)
    def writerows(self, rows):
        for row in rows:
            self.writerow(row)
class UnicodeDictWriter:
    """
    A CSV writer which will write dictionaries to CSV file "f",
    which is encoded in the given encoding.
    """
    def __init__(self, f, keys, dialect=csv.excel, encoding="utf-8", **kwds):
        # Redirect output to a queue
        self.queue = StringIO()
        self.writer = csv.writer(self.queue, dialect=dialect, **kwds)
        self.stream = f
        self.keys = keys
        self.encoder = codecs.getincrementalencoder(encoding)()
    def writeheader(self):
        header = OrderedDict()
        for key in self.keys:
            header.update({key:key})
        self.writerow(header)
    def writerow(self, row):
        self.writer.writerow([(row.get(key) if row.get(key) else "").encode("utf-8") for key in self.keys])
        # Fetch UTF-8 output from the queue ...
        data = self.queue.getvalue()
        data = data.decode("utf-8")
        # ... and reencode it into the target encoding
        data = self.encoder.encode(data)
        # write to the target stream
        self.stream.write(data)
        # empty queue
        self.queue.truncate(0)
    def writerows(self, rows):
        for row in rows:
            self.writerow(row)

### parse input file

if verbose:
  print "Reading " + (infile if not fromstdin else "from standard input") + "..."

inheader = None
invals = []
if not fromstdin:
  infd = codecs.open(infile, "rb", "utf8")
else:
  infd = infile

csvreader = UnicodeReader(infd, delimiter = "\t", quoting = csv.QUOTE_NONE, quotechar = '', escapechar = '\\')
for row in csvreader:
  if not inheader:
    inheader = row
    continue
  result = OrderedDict()
  for i in range(len(row)):
    result.update({inheader[i] : row[i]})
  invals.append(result)
if not fromstdin:
  infd.close()

### define DORA access

# prototype link:
# http://localhost:8080/solr/collection1/select?q=mods_identifier_local_s%3A10008&fq=PID%3Aeawag%5C%3A*&rows=20&fl=PID%2Cmods_identifier_local_*&wt=xml&indent=true

base_url = "http://localhost:8080/solr/collection1/select"
namespace = "eawag"
pidkey = "PID"
qfield = "mods_identifier_local_s" # later, set 'q' to qfield + ":" + RefID
fq = pidkey + ":" + namespace + "\:*" # restrict to namespace
fl_list = [] # define list of key-value-pairs to return
fl_list.append(pidkey)
fl_list.append(qfield)
fl = ", ".join(fl_list)
sort = pidkey + " desc" # get newest first
rows = 10 # actually, one should be enough!!
wt = 'json'
#indent = true # only needed if wt=xml


### get DORA PIDs

if verbose:
  print "Getting DORA PIDs..."

REFIDKEY = "ID"
DORAPIDKEY = "PID"
outheader = []
for key in inheader:
  outheader.append(key)
  if key == REFIDKEY:
    outheader.append(DORAPIDKEY)

outvals = []
for row in invals:
  result = OrderedDict()
  for key in row.keys():
    result.update({key : row[key]})
    if key == REFIDKEY and row[key]:
      result.update({DORAPIDKEY : None})
      payload = {'q' : qfield + ":" + row[key], 'fq' : fq, 'sort' : sort, 'fl' : fl, 'rows' : rows, 'wt' : wt} # add 'indent':indent if wt=xml
      r = requests.get(base_url, params=payload)
      if r.status_code != 200:
        print "Oops, something went wrong (got HTTP" + str(r.status_code) + ")!"
      else:
        rr = eval(r.text)
        response = rr["response"]
        if int(response["numFound"]) > 1:
          print "Warning: Found multiple records for RefID " + row[key] + "!"
          continue
        elif int(response["numFound"]) < 1 and verbose:
          print "Notice: Did not find any record for RefID " + row[key] + "."
          continue
        elif int(response["numFound"]) == 1:
          result[DORAPIDKEY] = response["docs"][0][pidkey]
  outvals.append(result)


### write result

if verbose:
  print "Writing resulting list to " + (outfile if not tostdout else "standard output") + "..."

if not tostdout:
  outfd = codecs.open(outfile, "wb", "utf8")
else:
  outfd = outfile

csvwriter = UnicodeDictWriter(outfd, outheader, delimiter = "\t", quoting = csv.QUOTE_NONE, quotechar = '', escapechar = '\\')

csvwriter.writeheader()
csvwriter.writerows(outvals)

if not tostdout:
  outfd.close()

