<!DOCTYPE html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>
</head>

# Eawag-Reporting 2016 – Mini-HowTo

0.	Preliminaries:
    1.	Create a working directory and a subdirectory “`2016_reports`” in it
    2.	Copy all the scripts in the working directory
    3.	The files to be saved in the following steps should also be copied in the working directory
    4.	The final reports will reside in the abovementioned subdirectory
1.	Get Journal Impact Factors
    1.	Go to https://jcr.incites.thomsonreuters.com/JCRJournalHomeAction.action?refineString=null&SrcApp=IC2LS&timeSpan=null&Init=Yes, customise Indicators (just enable everything; what we need is the ISSN and the JIF) and close the “window” (do not hit “save”, just the cross on the top right).
    2.	Hit the download symbol on the top right part of the page and save as csv to, say, “`20170103_JIF_Full_WoS_Export.csv`”.
    3.	Check for missing ISSNs: It turns out, that we are missing the ISSN-entry of “JOURNAL OF LOW FREQUENCY NOISE VIBRATION AND ACTIVE CONTROL”, which is 0263-0923 (see, e.g., http://www.multi-science.co.uk/lowfreq.htm or https://www.researchgate.net/journal/0263-0923_Journal_of_Low_Frequency_Noise_Vibration_and_Active_Control). An easy way to mend this, is to execute:
        ```
        cat 20170103_JIF_Full_WoS_Export.csv | sed '1 d; $ d;' | sed "`cat 20170103_JIF_Full_WoS_Export.csv | sed '1 d; $ d;' | sed -n '/^\([^,]*,\)\{3\},/! d; =;'` s/^\(\([^,]*,\)\{3\}\)/\10263-0923/" > 20170103_JIF_Full_WoS_Export.consolidated.csv
        ```
        This also removes the first and last lines. Note that this command works only in case this one particular ISSN is missing….
        **NOTE: If you save everything under the file names suggested in this document, you can simply run `./make_2016.sh` to perform all the shell commands from step 1 to step 9.**
2.	Get and prepare the old list for the 2015 reporting [optional]
    1.	Copy `\\`*2015_reporting_root*`\Reporting_2015_abgeschlossen\Rohdaten\Reporting_20160105_AllReferences.xlsx`
    2.	Use Excel (or LibreOffice Calc) to export the main sheet to tab-delimited text, after removing the two lines towards the end that separate the publications for reporting from those with publication dates before 2014 that will be ignored<font size="-1"><sup><a name="dagger_caller"></a>[&dagger;](#dagger_callee)</sup></font>. Save the file in, say, “`Reporting_20160105_AllReferences_all.txt`”
    3.	Finally, convert to utf8 & unix format using
        ```
        iconv -f windows-1252 -t utf-8 Reporting_20160105_AllReferences_all.txt | dos2unix > Reporting_20160105_AllReferences_all.utf8.txt
        ```
3.	Export this and last [optional] year’s publication list from RefWorks
    1.	Log in to RefWorks at http://www.refworks.com/refworks2/default.aspx?site=039241152255600000/RWWS5A549035/Eawag%20Publications
    2.	Go to “Search -> Advanced” and select “User 1” (reporting year) and 2016, resp. 2015
    3.	Hit “Create Bibliography”, and select “Tab Delimited with RefID*”  as output style and “Text” as output format
    4.	Since we have more than 500 publications in the export, the export might fail, in which case just try again (since 2016 has 577 publications and 2015 has 611, which are both not exaggerated, this should eventually work); once the “Completed” message is displayed in the lower right corner, right-click on the link and save the file as, say, “`2016_tab.txt`”, resp. “`2015_tab.txt`”
4.	Export this and last [optional] year’s report list from RefWorks
    1.	Log in to RefWorks at http://www.refworks.com/refworks2/default.aspx?site=039241152255600000/RWWS6A935776/Eawag%20Reports
    2.	Go to “Search -> Advanced” and select “User 1” (reporting year) and 2016, resp. 2015
    3.	Hit “Create Bibliography”, and select “Tab Delimited with RefID*”  as output style and “Text” as output format
    4.	Once the “Completed” message is displayed in the lower right corner, right-click on the link and save the file as, say, “`2016_reports_tab.txt`”, resp. “`2015_reports_tab.txt`”
5.	Convert the publication exports into UNIX format and nuke newlines within fields, as well as empty lines, using
    ```
    cat 2015_tab.txt | ./preprocess_exports.sh > 2015_tab.preprocessed.txt
    ```
    and
    ```
    cat 2016_tab.txt | ./preprocess_exports.sh > 2016_tab.preprocessed.txt
    ```
    Do the same thing for the reports:
    ```
    cat 2015_reports_tab.txt | ./preprocess_exports.sh > 2015_reports_tab.preprocessed.txt
    ```
    and
    ```
    cat 2016_reports_tab.txt | ./preprocess_exports.sh > 2016_reports_tab.preprocessed.txt
    ```
    (Note: the script “`preprocess_exports.sh`” depends on “`nuke_newlines.sed`”)
6.	Merge the reports with the “ordinary publication” using
    ```
    cat 2015_tab.txt > 2015_tab.all.txt && tail –n +2 2015_reports_tab.txt >> 2015_tab.all.txt
    ```
    and
    ```
    cat 2016_tab.txt > 2016_tab.all.txt && tail –n +2 2016_reports_tab.txt >> 2016_tab.all.txt
    ```
    (Note: The export-script distinguishes the two types of reports by looking if the RefID is above or below 10000)
7.	Compare the exports [optional]
    1.	Execute `PYTHONIOENCODING=utf8 python compare_old_and_new.py –v –l Reporting_20160105_AllReferences_all.utf8.txt –o 2015_tab.all.txt –n 2016_tab.all.txt 2> compare_old_and_new.log` and study the log-file “`compare_old_and_new.log`”
8.	Enrich the 2016 publications with their DORA-PIDs (if available)
    1.	Execute `PYTHONIOENCODING=utf8 python dora_lookup.py  -v < 2016_tab.all.txt > 2016_tab.all.withdorapids.txt`
9.	Finally, generate the reports
    1.	Create a subdirectory, say “`2016_reports`”, if not already done
    2.	Execute `PYTHONIOENCODING=utf8 python prepare_reports.py -v -p 2016_report -t "Report 2016" -o 2016_reports -y 2015 -j 20170103_JIF_Full_WoS-Export.consolidated.csv < 2016_tab.all.withdorapids.txt` 
10.	Now come the cosmetic parts:
    1.	In “`2016_reports`”, open each `docx`-file, decrease the text font, make the margins narrow, and display page numbers in the lower right corner, reducing their size, as well. Also, in the initial table, for all the numbers choose cell-alignment top-right.
    2.	Then, open in Excel (or LibreOffice Calc) the template `\\`*2016_reporting_root*`\templates\ArticlesTemplate_IF.xlsx` and, in the tab “Raw_Data”, insert, one by one, the data from `2016_report_`*???*`_articles.txt`, making sure to convert to utf-8 and set the text qualifier to “none” (Note: do not select all columns converting them to type text, as this makes filtering of, e.g., impact factors much harder). In the “Reporting” tab, drag the formula of the first table row down to the number of rows in “Raw_Data” plus three. Find the first row with “NR”, select the columns (CTRL+SHIFT+RIGHT from the left-most row), and add a border on top. Finally, check the title corresponds to the data (change A2, if necessary) and activate the filters in the header-line.
    3.	Prepare a similar document containing all references, so that people can more easily search and filter. More precisely, open a new xlsx-file, rename the tab to “Raw_Data” (consistency), and insert the data from `2016_report_`*???*`.txt` in the same way as above. Activate filters before saving.

<font size="-1"><sup><a name="dagger_callee"></a>[&dagger;](#dagger_caller)</sup>One could also use, e.g., Harald’s utility [`xlsxtocsv.exe`](https://github.com/eawag-rdm/xlsxtocsv) (after installation, you will find it in `C:\Users\`*your_user_name*`\AppData\Local`; you can run it, e.g., in a PowerShell). However, it exports to csv, and it’s a mess to convert this to tab delimited format via `sed`… ;)</font>

<br/>

> _This document is Copyright &copy; 2017 by d-r-p (Lib4RI) `<d-r-p@users.noreply.github.com>` and licensed under [CC&nbsp;BY&nbsp;4.0](https://creativecommons.org/licenses/by/4.0)._
