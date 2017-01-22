# eawag_reporting_2016

Our reporting workflow for Eawag publications in 2016

## Introduction

This is collection of tools used for the reporting of Eawag publications in 2016. The code and workflow are **obsolete**, insofar as the reporting will be done through our Islandora-based institutional repository [DORA Eawag](https://www.dora.lib4ri.ch/eawag) in the future. We placed them here so that parts can more easily be recycled.

**NOTE: This repository will not be maintained!**

## CAVEAT

THIS IS SERIOUSLY OUTDATED WORK!!! AT THE BEGINNING OF 2017, IT SEEMED TO HAVE DONE ITS JOB FOR US, BUT IT IS NOT ALWAYS CODED IN THE CORRECT WAY. DUE TO A SYSTEM CHANGE, THIS EVALUATION CODE WAS TO BE USED ONLY ONCE, AND IT WILL NOT EVER BE UPDATED. YOU SHOULD PROBABLY NOT USE THIS CODE YOURSELF, AS IT MIGHT NOT WORK FOR YOU OR EVEN BREAK YOUR SYSTEM (SEE ALSO 'LICENSE'). UNDER NO CIRCUMSTANCES WHATSOEVER ARE WE TO BE HELD LIABLE FOR ANYTHING. YOU HAVE BEEN WARNED.

## Requirements (@TODO: make the dependency list more explicit)

This software was successfully run on x86_64 GNU/Linux using
* [`python`](https://www.python.org) (2.7.9)
* the following additional packages from the repositories:
    - `iconv`
    - `python-pip`
    - `zip`
    - `unzip`
* the following python packages (install with `pip install`*package*
    - `docx`
* _possibly other tools not present in a default installation_

It is assumed, that all the scripts are placed together with the data in one working directory.

## Installation & Usage (@TODO: be more explicit)

1. Create a working directory.
2. Place all the scripts in the `scripts` subfolder in this working directory.
3. (**IMPORTANT**) In `scripts/dora_lookup.py`, replace "`localhost`" in the definition of `base_url` by the server running DORA's Solr instance.
4. Follow the instructions in `doc/2016_EawagReporting_MiniHowTo.md`.

## Output

If the instructions in `doc/2016_EawagReporting_MiniHowTo.md` are followed, the subdirectory `2016_reports` will contain a series of `docx` and `txt` files. Of each type, there will be one for each organisational unit, in addition to one containing the full list for Eawag (`*ALL*`). The `txt` files are tab separated lists, whereas each of the `docx` files contains a list of citations that can be assessed by the people responsible for the corresponding unit.

## Files (@TODO: be more explicit)

Note: Most scripts will accept the `-h` switch for short usage help.

### `.gitignore`

The file is such that the repository only contains the generic files, as well as itself. Specifically, it contains
```
/*

!/doc/
/doc/*
!/scripts/
/scripts/*
!/templates/
/templates/*

!/.gitignore
!/LICENSE
!/README.md
!/doc/2016_EawagReporting_MiniHowTo.md
!/scripts/amend_docx_styles.sh
!/scripts/compare_old_and_new.py
!/scripts/dora_lookup.py
!/scripts/make_2016.sh
!/scripts/nuke_newlines.sed
!/scripts/prepare_reports.py
!/scripts/preprocess_exports.sh
!/templates/ArticlesTemplate_IF.xlsx
```

### `LICENSE`

The license under which the toolchain is distributed:
```
Copyright (c) 2016, 2017 d-r-p (Lib4RI) <d-r-p@users.noreply.github.com>

Permission to use, copy, modify, and distribute this software for any
purpose with or without fee is hereby granted, provided that the above
copyright notice and this permission notice appear in all copies.

THE SOFTWARE IS PROVIDED "AS IS" AND THE AUTHOR DISCLAIMS ALL WARRANTIES
WITH REGARD TO THIS SOFTWARE INCLUDING ALL IMPLIED WARRANTIES OF
MERCHANTABILITY AND FITNESS. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR
ANY SPECIAL, DIRECT, INDIRECT, OR CONSEQUENTIAL DAMAGES OR ANY DAMAGES
WHATSOEVER RESULTING FROM LOSS OF USE, DATA OR PROFITS, WHETHER IN AN
ACTION OF CONTRACT, NEGLIGENCE OR OTHER TORTIOUS ACTION, ARISING OUT OF
OR IN CONNECTION WITH THE USE OR PERFORMANCE OF THIS SOFTWARE.
```

### `README.md`

This file...

### `doc/2016_EawagReporting_MiniHowTo.md`

The documentation of the workflow.

### `scripts/amend_docx_styles.sh`

Changes the styles for hyperlinks inside the resulting `docx` files.

### `scripts/compare_old_and_new.py`

Compares this year's data to last year's data (it's logic partially stems from the reporting 2015 having been done differently), then try to alert inconsistencies.

### `scripts/dora_lookup.py`

Fetches the DORA PID from the RefWorks ID. Processes a tab delimited file with first column the RefID. Returns the same file, but with a column containing the DORA PID, inserted after the first column.

**NOTE: It is absolutely indispensable to adjust the `base_url` variable, replacing "`localhost`" with the Solr server that is to be queried!**

### `scripts/make_2016.sh`

Crude script that runs the whole workflow. Assumes all files are named as in `doc/2016_EawagReporting_MiniHowTo.md`, and does not perform any check in this regard.

### `scripts/nuke_newlines.sed`

Removes newlines inside fields of tab separated data, as well as empty lines. Note that this could have been avoided with proper quoting.

### `scripts/prepare_reports.py`

Prepares the reports from pre-processed data.

### `scripts/preprocess_exports.sh`

Converts the raw data into a usable format.

### `templates/ArticlesTemplate_IF.xlsx`

A template to be filled with tab separated output data ("Raw Data" tab). In the "Reporting" tab, the most important information is re-displayed.

<br/>

> _This document is Copyright &copy; 2017 by d-r-p (Lib4RI) `<d-r-p@users.noreply.github.com>` and licensed under [CC&nbsp;BY&nbsp;4.0](https://creativecommons.org/licenses/by/4.0/)._