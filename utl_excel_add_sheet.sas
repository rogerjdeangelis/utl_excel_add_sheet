%let pgm=utl_excel_add_sheet;

Adding a sheet(tab) to a closed existing excel workbook

 If you have IML/R you can paste the code below into IML.
 I don't use 'proc export' so maybe that works?

 WORKING CODE
 ============

   females<-read_sas("d:/sd1/females.sas7bdat");             * get sas data to add to workbook;
   wb <- loadWorkbook("d:/xls/class.xlsx", create = FALSE);  * connect to existing workbook;
   writeWorksheet(wb,females,"females",startCol=1,header=T); * add female tab;

see
https://communities.sas.com/t5/Base-SAS-Programming/Import-tab-from-one-excel/m-p/371877


I already have a program that creates thr workbook and all its tabs.
However i have been asked to add a new tab. Ehich reada from other tabs.
Since its already created. Im wondering if i can impoer the tab
at the end or the creation job.



HAVE  (Workbook D:/XLS/HAVE.XLSX sheet MALES with males from sashelp.class
===================================================================================

SHEET ALL IN WORKBOOK D:/XLS/HAVE.XLSX

  +----------------------------------------------------------------+
  |     A      |    B       |     C      |    D       |    E       |
  +----------------------------------------------------------------+
1 |  NAME      |   SEX      |    AGE     |  HEIGHT    |  WEIGHT    |
  +------------+------------+------------+------------+------------+
2 | ALFRED     |    M       |    14      |    69      |  112.5     |
  +------------+------------+------------+------------+------------+
   ...
  +------------+------------+------------+------------+------------+
N | WILLIAM    |    F       |    15      |   66.5     |  112       |
  +------------+------------+------------+------------+------------+

[MALES]


WANT to add a second tab with FEMALES
=================================

SHEET MALES IN WORKBOOK D:/XLS/HAVE.XLSX

  +----------------------------------------------------------------+
  |     A      |    B       |     C      |    D       |    E       |
  +----------------------------------------------------------------+
1 |  NAME      |   SEX      |    AGE     |  HEIGHT    |  WEIGHT    |
  +------------+------------+------------+------------+------------+
2 | ALICE      |    F       |    14      |    69      |  112.5     |
  +------------+------------+------------+------------+------------+
   ...
  +------------+------------+------------+------------+------------+
N | WILMA      |    F       |    15      |   66.5     |  112       |
  +------------+------------+------------+------------+------------+

[FEMALES]


*                _                             _    _                 _
 _ __ ___   __ _| | _____  __      _____  _ __| | _| |__   ___   ___ | | __
| '_ ` _ \ / _` | |/ / _ \ \ \ /\ / / _ \| '__| |/ / '_ \ / _ \ / _ \| |/ /
| | | | | | (_| |   <  __/  \ V  V / (_) | |  |   <| |_) | (_) | (_) |   <
|_| |_| |_|\__,_|_|\_\___|   \_/\_/ \___/|_|  |_|\_\_.__/ \___/ \___/|_|\_\

;

%utlfkil(d:/xls/class.xlsx);
libname xel "d:/xls/class.xlsx" scan_text=no;
data xel.males;
  set sashelp.class(where=(sex='M'));
run;quit;
libname xel clear;

*                _                              _       _
 _ __ ___   __ _| | _____   ___  __ _ ___    __| | __ _| |_ __ _
| '_ ` _ \ / _` | |/ / _ \ / __|/ _` / __|  / _` |/ _` | __/ _` |
| | | | | | (_| |   <  __/ \__ \ (_| \__ \ | (_| | (_| | || (_| |
|_| |_| |_|\__,_|_|\_\___| |___/\__,_|___/  \__,_|\__,_|\__\__,_|

;

libname sd1 "d:/sd1";
options validvarname=upcase;
libname sd1 "d:/sd1";
data sd1.females;
  set sashelp.class(where=(sex='F'));
run;quit;

*          _       _   _
 ___  ___ | |_   _| |_(_) ___  _ __
/ __|/ _ \| | | | | __| |/ _ \| '_ \
\__ \ (_) | | |_| | |_| | (_) | | | |
|___/\___/|_|\__,_|\__|_|\___/|_| |_|

;

%utl_submit_r64('
source("c:/Program Files/R/R-3.3.2/etc/Rprofile.site",echo=T);
library(XLConnect);
library(haven);
females<-read_sas("d:/sd1/females.sas7bdat");
females;
wb <- loadWorkbook("d:/xls/class.xlsx", create = FALSE);
createSheet(wb, name = "females");
writeWorksheet(wb,females,"females",startCol=1,header=T);
saveWorkbook(wb);
');

