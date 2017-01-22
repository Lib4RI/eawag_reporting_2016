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
import docx # (install with 'pip install docx'; you may need python-pip and more)
from collections import OrderedDict
from cStringIO import StringIO

### some constants

STDIN   = sys.stdin  #'/dev/stdin'
STDOUT  = sys.stdout #'/dev/stdout'

allkey = u"ALL"
orgunits = [u"BB", u"DIR", u"ECO", u"ENG", u"ESS", u"FISHEC", u"KOM", u"OEKOTOX", u"SANDEC", u"SIAM", u"STAB", u"SURF", u"SWW", u"UCHEM", u"UMIK", u"UTOX", u"WUT"]
oualtnames = OrderedDict()
oualtnames.update({u"WUT" : [u"W+T"]})


### re-direct all output to stderr

oldstdout = sys.stdout
sys.stdout = sys.stderr

### parse the command line

usage = "Usage: %prog [-v] [-o outdir] [-p prefix] [-t title] [-j jiffile] [-y yearmin] [-i infile | infile]"
parser = OptionParser(usage)
parser.add_option("-v", "--verbose",
  action = "store_true", dest = "verbose", default = False,
  help = "Show what I'm doing [default=false]")
parser.add_option("-o", "--outdir", nargs = 1,
  action = "store", dest = "outdir", metavar = "outdir",
  help = "The output directory")
parser.add_option("-p", "--prefix", nargs = 1,
  action = "store", dest = "prefix", metavar = "prefix",
  help = "The prefix for the output files")
parser.add_option("-t", "--title", nargs = 1,
  action = "store", dest = "title", metavar = "title",
  help = "The title for the reference files")
parser.add_option("-j", "--jiffile", nargs = 1,
  action = "store", dest = "jiffile", metavar = "jiffile",
  help = "The JIF file (comma separated w/out 1st and last line)")
parser.add_option("-y", "--yearmin", nargs = 1,
  action = "store", dest = "yearmin", metavar = "yearmin",
  help = "The minimum publication year to be included in the report")
parser.add_option("-i", "--infile", nargs = 1,
  action = "store", dest = "infile", metavar = "infile",
  help = "The input file (tab delimited with RefID and DORA PID)")

(opts, args) = parser.parse_args()

verbose = opts.verbose

if verbose:
  print "Parsing command line..."

fromstdin = True

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

if not fromstdin and not os.access(infile, os.R_OK):
  parser.error("Cannot read from " + infile + "!")

jiffile = None
if opts.jiffile:
  jiffile = opts.jiffile

if jiffile and not os.access(jiffile, os.R_OK):
  parser.error("Cannot read from " + jiffile + "!")

yearmin = None
if opts.yearmin:
  try:
    yearmin = int(opts.yearmin)
  except ValueError:
    print "Warning: Need an integer for yearmin, got \"" + opts.yearmin + "\"; ignoring..."
    pass

outdir = "."
if opts.outdir:
  outdir = opts.outdir
if outdir[-1] == "/":
  outdir = outdir[:-1]
prefix = "report"
if opts.prefix:
  prefix = opts.prefix
title = "Report"
if opts.title:
  title = opts.title

outfiles = OrderedDict()
for key in ([allkey] + orgunits):
  outfiles.update({key : outdir + "/" + prefix + "_" + key + ".txt"})
articlesoutfiles = OrderedDict()
for key in outfiles:
  articlesoutfiles.update({key : outfiles[key][:-4] + "_articles.txt"})
statoutfiles = OrderedDict()
for key in outfiles:
  statoutfiles.update({key : outfiles[key][:-4] + ".docx"})

for outfile in outfiles.values() + articlesoutfiles.values() + statoutfiles.values():
  if not os.access(outfile, os.W_OK) and not (os.access(outdir, os.W_OK) and not os.access(outfile, os.F_OK)):
    parser.error("Cannot write to " + (outfile[2:] if outfile.find("./") == 0 else outfile) + "!")

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

### parse jif file, if it exists
jifdict = OrderedDict()
withissn = False
jifissnkey = "Issn"
jifjifkey = "Journal Impact Factor"
jiffjtkey = "Full Journal Title"
jifjatkey = "JCR Abbreviated Title"
if jiffile:
  if verbose:
    print "Parsing JIF file..."
  jiffd = codecs.open(jiffile, "rb", "utf8")
  csvreader = UnicodeReader(jiffd, delimiter = ",")
  jifheader = None
  for row in csvreader:
    if not jifheader:
      jifheader = row
      if jifissnkey in jifheader:
        withissn = True
      else:
        print "Warning: No ISSNs found in JIF file; consider re-exporting for better search results."
      continue
    result = OrderedDict()
    for i in range(len(row)):
      result.update({jifheader[i] : row[i].strip()})
    if withissn:
      jifdict.update({result[jifissnkey] : result})
    else:
      jifdict.update({result[jiffjtkey] : result})
  jiffd.close()

### define the jif-finder

def findjif(issn, jfull, jabbr):
  if not jiffile or ((not issn or issn == "") and (not jfull or jfull == "") and (not jabbr or jabbr == "")):
    return None
  # see http://dispatch.opac.d-nb.de/DB=1.1/SET=1/TTL=7/START_TEXT
  altissns = OrderedDict()
  altissns.update({'0148-0227' : ['2169-897X', '2169-9380', '2169-9313', '2169-9275', '2169-9097', '2169-9003', '2169-8953']}) # Journal of geophysical research: A-G
  key = jfull
  if withissn:
    key = issn
    for k in altissns:
      if key in altissns[k]:
        key = k
        break
  if key and key != "":
    if key in jifdict:
      return jifdict[key][jifjifkey]
  for val in jifdict.values():
    if issn and issn != "" and issn == val[jifissnkey]:
      return val[jifjifkey].strip()
    if jfull and jfull != "" and jfull.lower() == val[jiffjtkey].lower():
      return val[jifjifkey].strip()
    if jabbr and jabbr != "" and re.sub(".", "", jfull).lower() == val[jifjatkey].lower():
      return val[jifjifkey].strip()
  return None

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
    result.update({inheader[i] : row[i].strip()})
  invals.append(result)
if not fromstdin:
  infd.close()

### preparing the overall report

genrekey = "Reference Type"
titlekey = "Title Primary"
yrkey = "Pub Year"
oukey = "User 2"
excludegenres = ["Dissertation/Thesis"] #["Dissertation/Thesis", "Magazine Article", "Newspaper Article"]
def excludecorrigenda(row):
  exclude = False
  if row[titlekey].startswith("Corrig"):
    exclude = True
  return exclude
excludefilters = [excludecorrigenda]
def exclude(row):
  ryrkey = yrkey
  kwdskey = "Keywords"
  excludevals = excludegenres
  if row[genrekey] in excludevals:
    return True
  for f in excludefilters:
    if f(row):
      if verbose:
        print "Notice: Excluded reference RefId " + row["ID"] + " because of filters."
      return True
  try:
    yr = int(row[yrkey])
  except ValueError:
    print "Warning: Publication year \"" + row[yrkey] + "\" is not an integer in reference RefId " + row["ID"] + "!"
    pass
  else:
    if yearmin and yr < yearmin:
      if verbose:
        print "Notice: Excluded reference RefId " + row["ID"] + " because of year constraint."
      return True
    ounits = [k.strip() for k in row[oukey].split(",")]
    tmp = False
    for i in range(len(ounits)):
      for k in oualtnames:
        if ounits[i] in oualtnames[k]:
          ounits[i] = k
        break
    if any([k not in orgunits and not any([k in j for j in oualtnames.values()]) for k in ounits]):
      print "Warning: ambiguous value in field '" + oukey + "' in reference RefId " + row["ID"] + ": \"" + row[oukey] + "\""
    kwds = [k.strip() for k in row[kwdskey].split(";")]
    oukwds = []
    ouryrs = []
    for k in kwds:
      kk = re.sub("- Abteilung_([A-Z+]+_[0-9]+)", r'\1', k, flags = re.I) # case of the keyword with reporting year
      if kk == k: # no substitution took place
        kk = re.sub("- Abteilung_([A-Z+]+)", r'\1', kk, flags = re.I) # case of the keyword w/out reporting year
      if kk == k: # no substitution took place
        continue
      if kk.find("_") != -1:
        kku = re.sub("([A-Z+]+)_.*", r'\1', kk)
        kky = re.sub(".*_([0-9]+)", r'\1', kk)
      else:
        kku = kk
        kky = None
      oukwds.append(kku)
      if kky:
        ouryrs.append(kky)
    if len(set(ouryrs)) != 1:
      print "Warning: ambiguous or missing year in field '" + kwdskey + "' in reference RefId " + row["ID"]
    elif ouryrs[0] != row[ryrkey].strip():
      print "Warning: year mismatch in fields '" + kwdskey + "' and '" + ryrkey + "' in reference RefId " + row["ID"] + ": got \"" + ouryrs[0] + "\", resp. \"" + row[ryrkey] + "\""
    for i in range(len(oukwds)):
      for k in oualtnames:
        if oukwds[i] in oualtnames[k]:
          oukwds[i] = k
        break
    if any([k not in orgunits and not any([k in j for j in oualtnames.values()]) for k in oukwds]):
      print "Warning: ambiguous organisational unit in field '" + kwdskey + "' in reference RefId " + row["ID"]
    oulset = set(ounits) - set(oukwds)
    ourset = set(oukwds) - set(ounits)
    if len(oulset) != 0 or len(ourset) != 0:
      print "Warning: mismatching organisational units in fields '" + kwdskey + "' and '" + oukey + "' in reference RefId " + row["ID"] + ": got " + str(sorted(set(oukwds))) + ", resp. " + str(sorted(set(ounits)))
    if yearmin and yr < yearmin:
      if verbose:
        print "Notice: Excluded reference RefId " + row["ID"] + " because of year constraint."
      return True
  return False

if verbose:
  print "Preparing rows for export..."

outheader = []
jfullkey = "Periodical Full"
jabbrkey = "Periodical Abbrev"
issnkey = "ISSN/ISBN"
jifkey = "Impact Factor"
catkey = "User 4"
reviewedvals = ["SCI", "SCIE", "SSCI", "ESCI"]
rkey = "Reviewed"
nrkey = "Not reviewed"
for key in inheader:
  outheader.append(key)
  if jiffile and key == jabbrkey:
    outheader.append(jifkey)
outvals = []
ouarticles = OrderedDict()
articlegenres = ["Journal Article", "Magazine Article", "Newspaper Article"]
ouothers = OrderedDict()
othergenres = ["Book, Section", "Book, Whole", "Book, Edited", "Report", "Conference Proceedings"]
othergenres = sorted(othergenres)
allrarts = []
allnrarts = []
alloths = []
n = 0
for row in invals:
  if not exclude(row):
    result = OrderedDict()
    for i in range(len(outheader)):
      if outheader[i] == jifkey:
        val = findjif(row[issnkey], row[jfullkey], row[jabbrkey])
      else:
        val = row[outheader[i]]
      result.update({outheader[i] : val})
    if result[genrekey] not in articlegenres + othergenres:
      print "Warning: Uncategorised genre \"" + result[genrekey] + "\" in reference RefId " + result["ID"] + "!"
    if result[genrekey] in articlegenres:
      if result[catkey] in reviewedvals:
        allrarts.append(n)
      else:
        allnrarts.append(n)
    elif result[genrekey] in othergenres:
      alloths.append(n)
    for key in orgunits:
      keys = [key]
      if key in oualtnames:
        keys += oualtnames[key]
      if result[oukey] and any([result[oukey].find(k) != -1 for k in keys]):
        if result[genrekey] in articlegenres:
          if key not in ouarticles:
            ouarticles.update({key : OrderedDict()})
            ouarticles[key].update({rkey : []})
            ouarticles[key].update({nrkey : []})
          if result[catkey] in reviewedvals:
            ouarticles[key][rkey].append(n)
          else:
            ouarticles[key][nrkey].append(n)
        elif result[genrekey] in othergenres:
          if key not in ouothers:
            ouothers.update({key : []})
          ouothers[key].append(n)
    outvals.append(result)
    n += 1

if verbose:
  print "Sorting results..."

authorskey = "Authors, Primary"

# class from https://wiki.python.org/moin/HowTo/Sorting
class outvalssortedkey(object):
  def mycmp(self, a, b):
    global outvals
    r = 0 # we set this, so we can re-order more easily
    if r == 0:
      try:
        ia = int(outvals[int(a)]["ID"])
      except ValueError:
        ia = 0
        pass
      try:
        ib = int(outvals[int(b)]["ID"])
      except ValueError:
        ib = 0
        pass
      r = -(1 if ia >= 10000 and ib < 10000 else (-1 if ib >= 10000 and ia < 10000 else 0))
    if r == 0:
      r = cmp(outvals[int(a)][genrekey], outvals[int(b)][genrekey])
    if r == 0:
      try:
        ra = int(outvals[int(a)][yrkey])
      except ValueError:
        ra = 0
        pass
      try:
        rb = int(outvals[int(b)][yrkey])
      except ValueError:
        rb = 0
        pass
      r = -cmp(ra, rb)
    if r == 0:
      r = cmp(outvals[int(a)][authorskey], outvals[int(b)][authorskey])
    if r == 0:
      r = cmp(outvals[int(a)][titlekey], outvals[int(b)][titlekey])
    if r == 0:
      r = cmp(int(a), int(b))
    return r
  def __init__(self, obj, *args):
    self.obj = obj
  def __gt__(self, other):
    return self.mycmp(self.obj, other.obj) > 0
  def __lt__(self, other):
    return self.mycmp(self.obj, other.obj) < 0
  def __eq__(self, other):
    return self.mycmp(self.obj, other.obj) == 0
  def __ge__(self, other):
    return self.mycmp(self.obj, other.obj) >= 0
  def __le__(self, other):
    return self.mycmp(self.obj, other.obj) <= 0
  def __ne__(self, other):
    return self.mycmp(self.obj, other.obj) != 0

allrarts = sorted(allrarts, key = outvalssortedkey)
allnrarts = sorted(allnrarts, key = outvalssortedkey)
alloths = sorted(alloths, key = outvalssortedkey)

for key in orgunits:
  if key in ouarticles:
    ouarticles[key][rkey] = sorted(ouarticles[key][rkey], key = outvalssortedkey)
    ouarticles[key][nrkey] = sorted(ouarticles[key][nrkey], key = outvalssortedkey)
  if key in ouothers:
    ouothers[key] = sorted(ouothers[key], key = outvalssortedkey)

### exporting...

if verbose:
  print "Exporting to " + outfiles[allkey] + "..."

with codecs.open(outfiles[allkey], "wb", "utf8") as outfd:
  csvwriter = UnicodeDictWriter(outfd, outheader, delimiter = "\t", quoting = csv.QUOTE_NONE, quotechar = '', escapechar = '\\')
  csvwriter.writeheader()
  csvwriter.writerows([outvals[i] for i in allrarts])
  csvwriter.writerows([outvals[i] for i in allnrarts])
  csvwriter.writerows([outvals[i] for i in alloths])
  outfd.close()
  if verbose:
    print "Exported " + str(len(outvals)) + " items (" + str(len(allrarts)) + " reviewed, " + str(len(allnrarts)) + " non-reviewed articles and " + str(len(alloths)) + " other publication types)."

if verbose:
  print "Exporting articles to " + articlesoutfiles[allkey] + "..."

with codecs.open(articlesoutfiles[allkey], "wb", "utf8") as outfd:
  csvwriter = UnicodeDictWriter(outfd, outheader, delimiter = "\t", quoting = csv.QUOTE_NONE, quotechar = '', escapechar = '\\')
  csvwriter.writeheader()
  csvwriter.writerows([outvals[i] for i in allrarts])
  csvwriter.writerows([outvals[i] for i in allnrarts])
  outfd.close()
  if verbose:
    print "Exported " + str(len(allrarts + allnrarts)) + " articles."

if verbose:
  print "Exporting references to " + statoutfiles[allkey] + "..."

volumekey = "Volume"
issuekey = "Issue"
startpagekey = "Start Page"
otherpagekey = "Other Pages"
linkskey = "Links"
pidkey = "PID"
_addedhyperlinks = []
def add_hyperlink(par, url):
  global _addedhyperlinks
  _addedhyperlinks.append(url)
  hl = docx.makeelement('hyperlink', attrnsprefix = 'r', attributes = {'id': 'rId' + str(len(docx.relationshiplist()) + len(_addedhyperlinks))})
  run = docx.makeelement('r')
  rPr = docx.makeelement('rPr')
  rPr.append(docx.makeelement('rStyle', attributes = {'val' : 'Hyperlink'}))
  run.append(rPr)
  run.append(docx.makeelement('t', tagtext = url))
  hl.append(run)
  par.append(hl)
  return par

hyperlinkschema = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'
def makehyperlinkrelationshiplist():
  global _addedhyperlinks
  list_common = [hyperlinkschema]
  rlist = []
  for url in _addedhyperlinks:
    rlist.append(list_common + [url])
  _addedhyperlinks = []
  return rlist

def correctrelationships(rels):
  for r in rels.iter('Relationship'):
    if r.attrib['Type'] == hyperlinkschema:
      r.set('TargetMode', "External")
  return rels

sauthorskey = "Authors, Secondary"
stitlekey = "Title Secondary"
publisherkey = "Publisher"
placekey = "Place Of Publication"
doralinkprefix = "https://www.dora.lib4ri.ch/eawag/islandora/object/"
refworkslinkprefix = "http://www.refworks.com/refworks2/?site=039241152255600000%2fRWWS5A549035%2fEawag+Publications&rn="
refworksreportlinkprefix = "http://www.refworks.com/refshare2?site=039241152255600000/RWWS6A935776/Eawag%20Reports&rn="
def makeref(doc_body, row):
  refid = None
  try:
    refid = int(row["ID"])
  except ValueError:
    pass
  authors = None
  if row[authorskey] and row[authorskey] not in ["", " "]:
    authors = ", ".join([(", ".join([aa.strip() for aa in a.split(",")])) for a in row[authorskey].split(";")])
  cit_a = ((authors + " " if authors else "") + ("(" + row[yrkey] + ")" if row[yrkey] else "") + ". " if authors else "") + (row[titlekey] + ". " if row[titlekey] else "")
  cit_ab = ""
  cit_b = (row[jfullkey] if row[jfullkey] and row[jfullkey] not in ["", " "] else "")
  if cit_b == "":
    sauthors = None
    if row[sauthorskey] and row[sauthorskey] not in ["", " "]:
      sauthors = ", ".join([(", ".join([aa.strip() for aa in a.split(",")])) for a in row[sauthorskey].split(";")])
    cit_ab = ("In " + sauthors + " (Eds.)" if sauthors else "")
    cit_b = (row[stitlekey] if row[stitlekey] and row[stitlekey] not in ["", " "] else "")
    if cit_ab != "":
      cit_a += cit_ab + (", " if cit_b != "" else "")
  cit_c = (row[volumekey] if row[volumekey] and row[volumekey] not in ["", " "] else "") + ((" " if row[volumekey] and row[volumekey] not in ["", " "] else "") + "(" + row[issuekey] + ")" if row[issuekey] and row[issuekey] not in ["", " "] else "")
  cit_c = (cit_c + ", " if cit_c != "" else "") + (row[startpagekey] + ("-" + row[otherpagekey] if row[otherpagekey] and row[otherpagekey] not in ["", " "] else "") if row[startpagekey] and row[startpagekey] not in ["", " "] else "")
  cit_c = (cit_c + "." if cit_c != "" else "")
  cit_d = ""
  if not (row[jfullkey] and row[jfullkey] not in ["", " "]):
    cit_d = (row[placekey] if row[placekey] and row[placekey] not in ["", " "] else "") + (": " if row[placekey] and row[placekey] not in ["", " "] and row[publisherkey] and row[publisherkey] not in ["", " "] else "") + (row[publisherkey] if row[publisherkey] and row[publisherkey] not in ["", " "] else "")
  par_list = [(cit_a, '')]
  if cit_b != "":
    par_list.append((cit_b, 'i'))
  par_list.append((((", " if cit_b != "" or cit_ab != "" else "") + cit_c if cit_c != "" else "."), ''))
  if cit_d != "":
    par_list.append((" " + cit_d + ".", ''))
  p = docx.paragraph(par_list)
  # from https://github.com/python-openxml/python-docx/issues/74
  if row[linkskey] and row[linkskey] != "":
    run = docx.makeelement('r')
    rPr = docx.makeelement('rPr')
    run.append(rPr)
    t = docx.makeelement('t', tagtext = " ")
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    run.append(t)
    p.append(run)
    add_hyperlink(p, row[linkskey])
  if row[pidkey] and row[pidkey] != "":
    run = docx.makeelement('r')
    rPr = docx.makeelement('rPr')
    run.append(rPr)
    t = docx.makeelement('t', tagtext = " ")
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    run.append(t)
    p.append(run)
    add_hyperlink(p, doralinkprefix + row[pidkey])
  run = docx.makeelement('r')
  rPr = docx.makeelement('rPr')
  run.append(rPr)
  t = docx.makeelement('t', tagtext = " ")
  t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
  run.append(t)
  p.append(run)
  add_hyperlink(p, (refworksreportlinkprefix if (refid and refid < 10000) else refworkslinkprefix) + row["ID"])
  doc_body.append(p)
  return p

doc = docx.newdocument()
doc_body = None
for el in doc.iter():
  if el.tag == '{' + docx.nsprefixes['w'] + '}body':
    doc_body = el
    break
doc_body.append(docx.heading(title + " (" + str(len(outvals)) + ")", 1))
doc_body.append(docx.paragraph(["This report contains all publication types except for \"" + (excludegenres[0] if len(excludegenres) <= 1 else "\"; \"".join([excludegenres[k] for k in range(len(excludegenres)-1)]) + "\" and \"" + excludegenres[len(excludegenres)-1]) + "\"" + (", as well as publications published before " + str(yearmin) if yearmin else "") + ".", '']))
doc_body.append(docx.heading("Statistics", 2))
stats_table = []
stats_table.append(["", "Articles" + u" \u2013 " + rkey, "Articles" + u" \u2013 " + nrkey, "Others", "TOTAL"])
for key in orgunits:
  n_rarts = len(ouarticles[key][rkey]if key in ouarticles else [])
  n_nrarts = len(ouarticles[key][nrkey] if key in ouarticles else [])
  n_oths = len(ouothers[key] if key in ouothers else [])
  n_tot = n_rarts + n_nrarts + n_oths
  stats_table.append([key, str(n_rarts), str(n_nrarts), str(n_oths), str(n_tot)])
stats_table.append(["TOTAL Eawag", str(len(allrarts)), str(len(allnrarts)), str(len(alloths)), str(len(outvals))])
doc_body.append(docx.table(stats_table, colw = [2000, 1700, 1700, 1700, 1700], cwunit = 'dxa'))
doc_body.append(docx.paragraph(["Note: Articles include the publication types \"" + "\", \"".join([articlegenres[k] for k in range(len(articlegenres)-1)]) + "\" and \"" + articlegenres[len(articlegenres)-1] + "\", whereas Others include the publication types \"" + "\", \"".join([othergenres[k] for k in range(len(othergenres)-1)]) + "\" and \"" + othergenres[len(othergenres)-1] + "\".", '']))
doc_body.append(docx.heading(" / ".join(articlegenres) + u" \u2013 " + rkey + " (" + str(len(allrarts)) + ")", 2))
for n in allrarts:
  makeref(doc_body, outvals[n])
doc_body.append(docx.heading(" / ".join(articlegenres) + u" \u2013 " + nrkey + " (" + str(len(allnrarts)) + ")", 2))
for n in allnrarts:
  makeref(doc_body, outvals[n])
doc_body.append(docx.heading(" / ".join(othergenres) + " (" + str(len(alloths)) + ")", 2))
for n in alloths:
  makeref(doc_body, outvals[n])
doc_rels = docx.wordrelationships(docx.relationshiplist() + makehyperlinkrelationshiplist())
doc_rels = correctrelationships(doc_rels)
docx.savedocx(doc, docx.coreproperties(title, "", "prepare_reports.py", ""), docx.appproperties(), docx.contenttypes(), docx.websettings(), doc_rels, statoutfiles[allkey])

for key in orgunits:
  if verbose:
    print "Exporting to " + outfiles[key] + "..."

  with codecs.open(outfiles[key], "wb", "utf8") as outfd:
    csvwriter = UnicodeDictWriter(outfd, outheader, delimiter = "\t", quoting = csv.QUOTE_NONE, quotechar = '', escapechar = '\\')
    csvwriter.writeheader()
    rarts = ouarticles[key][rkey] if key in ouarticles else []
    nrarts = ouarticles[key][nrkey] if key in ouarticles else []
    oths = ouothers[key] if key in ouothers else []
    if rarts != []:
      csvwriter.writerows([outvals[k] for k in rarts])
    if nrarts != []:
      csvwriter.writerows([outvals[k] for k in nrarts])
    if oths != []:
      csvwriter.writerows([outvals[k] for k in oths])
    outfd.close()
    if verbose:
      print "Exported " + str(len(rarts) + len(nrarts) + len(oths)) + " items (" + str(len(rarts)) + " reviewed, " + str(len(nrarts)) + " non-reviewed articles and " + str(len(oths)) + " other publication types)."
  if verbose:
    print "Exporting articles to " + articlesoutfiles[key] + "..."
  with codecs.open(articlesoutfiles[key], "wb", "utf8") as outfd:
    csvwriter = UnicodeDictWriter(outfd, outheader, delimiter = "\t", quoting = csv.QUOTE_NONE, quotechar = '', escapechar = '\\')
    csvwriter.writeheader()
    if rarts != []:
      csvwriter.writerows([outvals[i] for i in rarts])
    if nrarts != []:
      csvwriter.writerows([outvals[i] for i in nrarts])
    outfd.close()
    if verbose:
      print "Exported " + str(len(rarts + nrarts)) + " articles."
  if verbose:
    print "Exporting references to " + statoutfiles[key] + "..."
  doc = docx.newdocument()
  doc_body = None
  for el in doc.iter():
    if el.tag == '{' + docx.nsprefixes['w'] + '}body':
      doc_body = el
      break
  doc_body.append(docx.heading(title + u" \u2013 " + key + " (" + str(len(rarts) + len(nrarts) + len(oths)) + ")", 1))
  doc_body.append(docx.paragraph(["This report contains all publication types except for \"" + (excludegenres[0] if len(excludegenres) <= 1 else "\"; \"".join([excludegenres[k] for k in range(len(excludegenres)-1)]) + "\" and \"" + excludegenres[len(excludegenres)-1]) + "\"" + (", as well as publications published before " + str(yearmin) if yearmin else "") + ".", '']))
  doc_body.append(docx.heading("Statistics", 2))
  stats_table = []
  stats_table.append([" / ".join(articlegenres) + u" \u2013 " + rkey, " / ".join(articlegenres) + u" \u2013 " + nrkey, "Others", "TOTAL"])
  n_rarts = len(rarts)
  n_nrarts = len(nrarts)
  n_oths = len(oths)
  n_tot = n_rarts + n_nrarts + n_oths
  stats_table.append([str(n_rarts), str(n_nrarts), str(n_oths), str(n_tot)])
  doc_body.append(docx.table(stats_table, colw = [2200, 2200, 2200, 2200], cwunit='dxa'))
  doc_body.append(docx.paragraph(["Note: Others include the publication types \"" + "\", \"".join([othergenres[k] for k in range(len(othergenres)-1)]) + "\" and \"" + othergenres[len(othergenres)-1] + "\"", '']))
  doc_body.append(docx.heading(" / ".join(articlegenres) + u" \u2013 " + rkey + " (" + str(len(ouarticles[key][rkey] if key in ouarticles else [])) + ")", 2))
  for n in (ouarticles[key][rkey] if key in ouarticles else []):
    makeref(doc_body, outvals[n])
  doc_body.append(docx.heading(" / ".join(articlegenres) + u" \u2013 " + nrkey + " (" + str(len(ouarticles[key][nrkey] if key in ouarticles else [])) + ")", 2))
  for n in (ouarticles[key][nrkey] if key in ouarticles else []):
    makeref(doc_body, outvals[n])
  doc_body.append(docx.heading(" / ".join(othergenres) + " (" + str(len(ouothers[key] if key in ouothers else [])) + ")", 2))
  for n in (ouothers[key] if key in ouothers else []):
    makeref(doc_body, outvals[n])
  doc_rels = docx.wordrelationships(docx.relationshiplist() + makehyperlinkrelationshiplist())
  doc_rels = correctrelationships(doc_rels)
  docx.savedocx(doc, docx.coreproperties(title + u" \u2013 " + key, "", "prepare_reports.py", ""), docx.appproperties(), docx.contenttypes(), docx.websettings(), doc_rels, statoutfiles[key])

