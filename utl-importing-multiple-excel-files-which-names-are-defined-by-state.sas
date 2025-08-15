%let pgm=utl-importing-multiple-excel-files-which-names-are-defined-by-state;

%stop_submission;

Importing multiple excel files which names are defined by state

github
https://tinyurl.com/43vc9n3z
https://github.com/rogerjdeangelis/utl-importing-multiple-excel-files-which-names-are-defined-by-state

sas communities
https://tinyurl.com/hzjcxuer
https://communities.sas.com/t5/SAS-Procedures/Importing-multiple-excel-files-which-names-are-defined-by-state/m-p/765622#M80974

SOAPBOX ON
  I assume the layout of all the workbooks is the same,
  otherwise you would need up to 1000 programs to process the data.
  I also assume you have meta data with the workbook names and sheets.
SOAPBOX OFF

/******************************************************************************************************************************/
/* INPUT                                   | PROCESS                                   | OUTPUT (THREE WORKBOOKS)             */
/* =====                                   | =======                                   | ======                               */
/* SD1.META                                | proc datasets lib=sd1                     |                 SOURCE_              */
/*                                         |   nolist nodetails;                       | #  NAME  SEX AGE SHEET SOURCE_FILE   */
/* WORKBOOK           WORKSHEET            |  delete want;                             |                                      */
/*                                         | run;quit;                                 | 1 Alfred  M  14 S98ZPAL 98zpAL.xlsx  */
/* d:/xls/98zpAL.xlsx 98zpAL               |                                           | 2 Alice   F  13 S98ZPAL 98zpAL.xlsx  */
/* d:/xls/98zpFL.xlsx 98zpFL               | %utl_rbeginx;                             | 3 Barbara F  13 S98ZPAL 98zpAL.xlsx  */
/* d:/xls/98zpTX.xlsx 98zpTX               | parmcards4;                               | 4 Carol   F  14 S98ZPAL 98zpAL.xlsx  */
/*                                         | library(haven)                            | 5 Henry   M  14 S98ZPAL 98zpAL.xlsx  */
/* One of the three workbooks              | library(readxl)                           | 6 James   M  12 S98ZPAL 98zpAL.xlsx  */
/* d:/xls/98zpAL.xlsx                      | library(dplyr)                            |                                      */
/*                                         | source("c:/oto/fn_tosas9x.R")             | 1 Alfred  M  14 S98ZPFL 98zpFL.xlsx  */
/* --------------------------+             | meta<-read_sas("d:/sd1/meta.sas7bdat")    | 2 Alice   F  13 S98ZPFL 98zpFL.xlsx  */
/* | A1| fx       |DAYNUM    |             | meta                                      | 3 Barbara F  13 S98ZPFL 98zpFL.xlsx  */
/* -----------------------------------+    | list_of_dfs <-                            | 4 Carol   F  14 S98ZPFL 98zpFL.xlsx  */
/* [_] |    A     |    B    |    C    |    |  lapply(seq_len(nrow(meta)),function(i) { | 5 Henry   M  14 S98ZPFL 98zpFL.xlsx  */
/* -----------------------------------|    |   df <- read_excel(meta$WORKBOOK[i]) %>%  | 6 James   M  12 S98ZPFL 98zpFL.xlsx  */
/*  1  | NAME     |   SEX   |   AGE   |    |    mutate(                                |                                      */
/*  -- |----------+---------+---------|    |     source_worksheet=meta$WORKSHEET[i],   | 1 Alfred  M  14 S98ZPTX 98zpTX.xlsx  */
/*  2  |  Alfred  | M       | 14      |    |     source_file=basename(meta$WORKBOOK[i])| 2 Alice   F  13 S98ZPTX 98zpTX.xlsx  */
/*  -- |----------+---------+---------|    |    )                                      | 3 Barbara F  13 S98ZPTX 98zpTX.xlsx  */
/*  3  |  Alice   | F       | 13      |    |   return(df)                              | 4 Carol   F  14 S98ZPTX 98zpTX.xlsx  */
/*  -- |----------+---------+---------|    | })                                        | 5 Henry   M  14 S98ZPTX 98zpTX.xlsx  */
/*  4  |  Barbara | F       | 13      |    | combined_df <- bind_rows(list_of_dfs)     | 6 James   M  12 S98ZPTX 98zpTX.xlsx  */
/*  -- |----------+---------+---------|    | fn_tosas9x(                               |                                      */
/*  5  |  Carol   | F       | 14      |    |       inp    = combined_df                |                                      */
/*  -- |----------+---------+---------|    |      ,outlib ="d:/sd1/"                   |                                      */
/*  6  |  Henry   | M       | 14      |    |      ,outdsn ="want"                      |                                      */
/*  -- |----------+---------+---------|    |      )                                    |                                      */
/*  7  |  James   | M       | 12      |    | ;;;;                                      |                                      */
/*  -- |----------+---------+---------|    | %utl_rendx;                               |                                      */
/*  [98zpAL]                               |                                           |                                      */
/*                                         |                                           |                                      */
/*                                         | proc print data=sd1.want;                 |                                      */
/* options validvarname=upcase;            | run;quit;                                 |                                      */
/* libname sd1 "d:/sd1";                   |                                           |                                      */
/* data sd1.class;                         |                                           |                                      */
/*   input                                 |                                           |                                      */
/*     name$                               |                                           |                                      */
/*     sex$ age;                           |                                           |                                      */
/* cards4;                                 |                                           |                                      */
/* Alfred  M 14                            |                                           |                                      */
/* Alice   F 13                            |                                           |                                      */
/* Barbara F 13                            |                                           |                                      */
/* Carol   F 14                            |                                           |                                      */
/* Henry   M 14                            |                                           |                                      */
/* James   M 12                            |                                           |                                      */
/* ;;;;                                    |                                           |                                      */
/* run;quit;                               |                                           |                                      */
/*                                         |                                           |                                      */
/*                                         |                                           |                                      */
/* %deletesasmacn;                         |                                           |                                      */
/* %symdel zst wbs / nowarn;               |                                           |                                      */
/*                                         |                                           |                                      */
/* options validvarname=upcase;            |                                           |                                      */
/* libname sd1 "d:/sd1";                   |                                           |                                      */
/* data sd1.meta;                          |                                           |                                      */
/*  length workbook $255;                  |                                           |                                      */
/*  input worksheet$;                      |                                           |                                      */
/*  call symputx('worksheet',worksheet);   |                                           |                                      */
/*                                         |                                           |                                      */
/*  workbook=cats('d:/xls/'                |                                           |                                      */
/*     ,worksheet,'.xlsx');                |                                           |                                      */
/*  call symputx('workbook',workbook);     |                                           |                                      */
/*  rc=dosubl('                            |                                           |                                      */
/*    %utlfkil(&worksheet);                |                                           |                                      */
/*    ods excel file="&workbook"           |                                           |                                      */
/*      options(sheet_name="&worksheet");  |                                           |                                      */
/*    proc print data=sd1.class;           |                                           |                                      */
/*    run;quit;                            |                                           |                                      */
/*    ods excel close;                     |                                           |                                      */
/*    ');                                  |                                           |                                      */
/*  worksheet=cats('S',upcase(worksheet)); |                                           |                                      */
/* drop rc;                                |                                           |                                      */
/* cards4;                                 |                                           |                                      */
/* 98zpAL                                  |                                           |                                      */
/* 98zpFL                                  |                                           |                                      */
/* 98zpTX                                  |                                           |                                      */
/* ;;;;                                    |                                           |                                      */
/* run;quit;                               |                                           |                                      */
/******************************************************************************************************************************/

/*                   _
(_)_ __  _ __  _   _| |_
| | `_ \| `_ \| | | | __|
| | | | | |_) | |_| | |_
|_|_| |_| .__/ \__,_|\__|
        |_|
*/

options validvarname=upcase;
libname sd1 "d:/sd1";
data sd1.class;
  input
    name$
    sex$ age;
cards4;
Alfred  M 14
Alice   F 13
Barbara F 13
Carol   F 14
Henry   M 14
James   M 12
;;;;
run;quit;


%deletesasmacn;
%symdel zst wbs / nowarn;

options validvarname=upcase;
libname sd1 "d:/sd1";
data sd1.meta;
 length workbook $255;
 input worksheet$;
 call symputx('worksheet',worksheet);

 workbook=cats('d:/xls/'
    ,worksheet,'.xlsx');
 call symputx('workbook',workbook);
 rc=dosubl('
   %utlfkil(&worksheet);
   ods excel file="&workbook"
     options(sheet_name="&worksheet");
   proc print data=sd1.class;
   run;quit;
   ods excel close;
   ');
 worksheet=cats('S',upcase(worksheet));
drop rc;
cards4;
98zpAL
98zpFL
98zpTX
;;;;
run;quit;

/******************************************************************************************************************************/
/* INPUT                                                                                                                      */
/* =====                                                                                                                      */
/* SD1.META                                                                                                                   */
/*                                                                                                                            */
/* WORKBOOK           WORKSHEET                                                                                               */
/*                                                                                                                            */
/* d:/xls/98zpAL.xlsx 98zpAL                                                                                                  */
/* d:/xls/98zpFL.xlsx 98zpFL                                                                                                  */
/* d:/xls/98zpTX.xlsx 98zpTX                                                                                                  */
/*                                                                                                                            */
/* One of the three workbooks                                                                                                 */
/* d:/xls/98zpAL.xlsx                                                                                                         */
/*                                                                                                                            */
/* --------------------------+                                                                                                */
/* | A1| fx       |DAYNUM    |                                                                                                */
/* -----------------------------------+                                                                                       */
/* [_] |    A     |    B    |    C    |                                                                                       */
/* -----------------------------------|                                                                                       */
/*  1  | NAME     |   SEX   |   AGE   |                                                                                       */
/*  -- |----------+---------+---------|                                                                                       */
/*  2  |  Alfred  | M       | 14      |                                                                                       */
/*  -- |----------+---------+---------|                                                                                       */
/*  3  |  Alice   | F       | 13      |                                                                                       */
/*  -- |----------+---------+---------|                                                                                       */
/*  4  |  Barbara | F       | 13      |                                                                                       */
/*  -- |----------+---------+---------|                                                                                       */
/*  5  |  Carol   | F       | 14      |                                                                                       */
/*  -- |----------+---------+---------|                                                                                       */
/*  6  |  Henry   | M       | 14      |                                                                                       */
/*  -- |----------+---------+---------|                                                                                       */
/*  7  |  James   | M       | 12      |                                                                                       */
/*  -- |----------+---------+---------|                                                                                       */
/*  [98zpAL]                                                                                                                  */
/******************************************************************************************************************************/

/*
 _ __  _ __ ___   ___ ___  ___ ___
| `_ \| `__/ _ \ / __/ _ \/ __/ __|
| |_) | | | (_) | (_|  __/\__ \__ \
| .__/|_|  \___/ \___\___||___/___/
|_|
*/

proc datasets lib=sd1
  nolist nodetails;
 delete want;
run;quit;

%utl_rbeginx;
parmcards4;
library(haven)
library(readxl)
library(dplyr)
source("c:/oto/fn_tosas9x.R")
meta<-read_sas("d:/sd1/meta.sas7bdat")
meta
list_of_dfs <-
 lapply(seq_len(nrow(meta)),function(i) {
  df <- read_excel(meta$WORKBOOK[i]) %>%
   mutate(
    source_worksheet=meta$WORKSHEET[i],
    source_file=basename(meta$WORKBOOK[i])
   )
  return(df)
})
combined_df <- bind_rows(list_of_dfs)
combined_df
fn_tosas9x(
      inp    = combined_df
     ,outlib ="d:/sd1/"
     ,outdsn ="want"
     )
;;;;
%utl_rendx;


proc print data=sd1.want;
run;quit;

/**************************************************************************************************************************/
/* R                                                        | SAS                                                         */
/*> combined_df                                             |                                  SOURCE_                    */
/*# A tibble: 18 Ã— 6                                       | OBS     NAME      SEX    AGE    WORKSHEET    SOURCE_FILE    */
/*     Obs NAME    SEX     AGE source_worksheet source_file |                                                             */
/*   <dbl> <chr>   <chr> <dbl> <chr>            <chr>       |  1     Alfred      M      14     S98ZPAL     98zpAL.xlsx    */
/* 1     1 Alfred  M        14 S98ZPAL          98zpAL.xlsx |  2     Alice       F      13     S98ZPAL     98zpAL.xlsx    */
/* 2     2 Alice   F        13 S98ZPAL          98zpAL.xlsx |  3     Barbara     F      13     S98ZPAL     98zpAL.xlsx    */
/* 3     3 Barbara F        13 S98ZPAL          98zpAL.xlsx |  4     Carol       F      14     S98ZPAL     98zpAL.xlsx    */
/* 4     4 Carol   F        14 S98ZPAL          98zpAL.xlsx |  5     Henry       M      14     S98ZPAL     98zpAL.xlsx    */
/* 5     5 Henry   M        14 S98ZPAL          98zpAL.xlsx |  6     James       M      12     S98ZPAL     98zpAL.xlsx    */
/* 6     6 James   M        12 S98ZPAL          98zpAL.xlsx |  1     Alfred      M      14     S98ZPFL     98zpFL.xlsx    */
/* 7     1 Alfred  M        14 S98ZPFL          98zpFL.xlsx |  2     Alice       F      13     S98ZPFL     98zpFL.xlsx    */
/* 8     2 Alice   F        13 S98ZPFL          98zpFL.xlsx |  3     Barbara     F      13     S98ZPFL     98zpFL.xlsx    */
/* 9     3 Barbara F        13 S98ZPFL          98zpFL.xlsx |  4     Carol       F      14     S98ZPFL     98zpFL.xlsx    */
/*10     4 Carol   F        14 S98ZPFL          98zpFL.xlsx |  5     Henry       M      14     S98ZPFL     98zpFL.xlsx    */
/*11     5 Henry   M        14 S98ZPFL          98zpFL.xlsx |  6     James       M      12     S98ZPFL     98zpFL.xlsx    */
/*12     6 James   M        12 S98ZPFL          98zpFL.xlsx |  1     Alfred      M      14     S98ZPTX     98zpTX.xlsx    */
/*13     1 Alfred  M        14 S98ZPTX          98zpTX.xlsx |  2     Alice       F      13     S98ZPTX     98zpTX.xlsx    */
/*14     2 Alice   F        13 S98ZPTX          98zpTX.xlsx |  3     Barbara     F      13     S98ZPTX     98zpTX.xlsx    */
/*15     3 Barbara F        13 S98ZPTX          98zpTX.xlsx |  4     Carol       F      14     S98ZPTX     98zpTX.xlsx    */
/*16     4 Carol   F        14 S98ZPTX          98zpTX.xlsx |  5     Henry       M      14     S98ZPTX     98zpTX.xlsx    */
/*17     5 Henry   M        14 S98ZPTX          98zpTX.xlsx |  6     James       M      12     S98ZPTX     98zpTX.xlsx    */
/*18     6 James   M        12 S98ZPTX          98zpTX.xlsx |                                                             */
/**************************************************************************************************************************/

/*              _
  ___ _ __   __| |
 / _ \ `_ \ / _` |
|  __/ | | | (_| |
 \___|_| |_|\__,_|

*/
