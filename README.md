# utl-altair-slc-language6-rexcel-read-write-text-files-summarize-and-looping-with-sqlite
Altair slc language6 rexcel read write text files summarize and looping with sqlite
    %let pgm=utl-altair-slc-language6-rexcel-read-write-text-files-summarize-and-looping-with-sqlite;

    %stop_submission;

    Altair slc language6 rexcel read write text files summarize and looping with sqlite

    SOAPBOX ON
    I don't think it is possible to add sheets to a CLOSED excel workbook using sas.
    Alos I don't thin sas can preserve the style of yhe workbook.
    I even doublt you can do it with sas access?
    r excel and python (with the code below) can add sheets to a CLOSED workbook
    SOAPBOX OFF

    Note: None of the process queries are supported by the slc 'proc sql' or sas 'proc sql'
    Note: Character strings are defaulting to 1024 bytes, to fix this you need to define
    the table attributes explicitly in sqlite. Macro utl_optlenpos will optimize the lengths of all variables.

    Too long to post on a list, see github
    https://github.com/rogerjdeangelis/utl-altair-slc-language6-rexcel-read-write-text-files-summarize-and-looping-with-sqlite


    Note: Note sqlite has csv read and write commands, however I want to present general txt processing in sql.
          Keep in mind that I am demonstrating that the exact sqlite queries work in the 8 languages below.
          In addition you can use the exact same queries in any language or operating system that supports odbc.

    related repos (see for information of installing odbc and sqlite)

    https://github.com/rogerjdeangelis/utl-altair-slc-sqlite-cheat-sheet
    https://github.com/rogerjdeangelis/utl-altair-slc-language1-drop-down-to-open-source-spss-and-execute-postgresql-query
    https://github.com/rogerjdeangelis/utl-altair-slc-language2-drop-down-to-open-source-matlab-and-execute-sqlite-with-extensions
    https://github.com/rogerjdeangelis/utl-altair-slc-language3-proc-r-read-write-csv-files-and-summarize-and-looping-with-sqlite
    https://github.com/rogerjdeangelis/utl-altair-slc-language4-proc-python-read-write-csv-files-and-summarize-and-looping-with-sqlite
    https://github.com/rogerjdeangelis/utl-altair-slc-native-language5-odbc-read-write-text-files-summarize-and-looping-with-sqlite
    https://github.com/rogerjdeangelis/utl-altair-slc-language6-rexcel-read-write-text-files-summarize-and-looping-with-sqlite

    PROBLEM


        d:/xls/class_export.xlsx  sheet='have' (we will add sheets)

        ORIGINAL CLOSED WORKBOOK (preserved style)
        -------------------------------------

        -------------------------+
        | A1| fx       | NAME    |
        ---------------------------------------------------------------------------+
        [_] |    A     |    B    |    C    |   D     |          E                  |
        ---------------------------------------------------------------------------|
         1  | NAME     |   SEX   |   AGE   | HEIGHT  |        LINE                 |
         -- |----------+---------+---------+---------+-----------------------------|
         2  |  Alfred  | M       | 14      | 69      | This is the 1st line        |
         -- |----------+---------+---------+---------+-----------------------------|
         3  |  Alice   | F       | 13      | 56.5    | This is the 2nd line        |
         -- |----------+---------+---------+---------+-----------------------------|
         4  |  Barbara | F       | 13      | 65.3    | This is the 3rd line        |
         -- |----------+---------+---------+---------+-----------------------------|
         5  |  Carol   | F       | 14      | 62.8    | This is the 4th line        |
         -- |----------+---------+---------+---------+-----------------------------|
         6  |  Henry   | M       | 14      | 63.5    | This is the 5th line        |
         -- |----------+---------+---------+---------------------------------------|
        [HAVE]


        CREATE TEXT FILE USING ONLY SQLITE WRITEFILE
        ---------------------------------------------

         d:/txt/class_export.txt

         This is the 1st line
         This is the 2nd line
         This is the 3rd line
         This is the 4th line
         This is the 5th line


        ADD SHEET LINE USING ONLY SQLITE READFILE (TO THE CLOSED WORKBOOK)
        ------------------------------------------------------------------

        -------------------------+
        | A1| fx       | LINE    |
        ----------------------------------+
        [_] |           A                 |
        ----------------------------------|
         1  |        LINE                 |
         -- |-----------------------------|
         2  | This is the 1st line        |
         -- |-----------------------------|
         3  | This is the 2nd line        |
         -- |-----------------------------|
         4  | This is the 3rd line        |
         -- |-----------------------------|
         5  | This is the 4th line        |
         -- |-----------------------------|
         6  | This is the 5th line        |
         -- |-----------------------------+
        [LINE]


        ADD SHEET AVGS
        ---------------

        -----------------------------------+
        | A1| fx                 |NAMES    |
        ------------------------------------------------------------
        [_] |    A               |    B    |    C    |      D      |
        ------------------------------------------------------------
         1  | NAMES              |   SEX   |AVG_AGE  |AVG_ HEIGHT  |
         -- |--------------------+---------+---------+--------------
         2  | Alice,Barbara,Caro l       F |   13.3  |   61.5      |
         -- |----------+---------+---------+---------+--------------
         3  | Alfred,Henry       |       M |   14    |   66.3      |
         -- |--------------------+---------+---------+--------------
        [AVGS]


     PROCESS

       1 slc create sqlite table 'have' with 5 lines (sentences) and colu

         d:/sqlite/mysqlite.db table have

         line                        name        sex    age    height

         This is the 1st line        Alfred       M      14     69.0
         This is the 2nd line        Alice        F      13     56.5
         This is the 3rd line        Barbara      F      13     65.3
         This is the 4th line        Carol        F      14     62.8
         This is the 5th line        Henry        M      14     63.5

         create workbook and sheet have
         d:/xlsx/class_export.xlsx sheet have
         we will add sheets to this closed workbook using r

         line                        name        sex    age    height

         This is the 1st line        Alfr ed       M      14     69.0
         This is the 2nd line        Alice        F      13     56.5
         This is the 3rd line        Barbara      F      13     65.3
         This is the 4th line        Carol        F      14     62.8
         This is the 5th line        Henry        M      14     63.5

      2  slc summarize sqlite table

         select sex avg(age) avg(height) from have group by sex


      3  slc convert sqlite table 'have' to text file using sqlite 'writefile' command

         d:/txt/have.txt

         This is the 1st line
         This is the 2nd line
         This is the 3rd line

      4  slc convert text file 'd:/txt/have.txt' to slc dataframe using sqlite readfile

          text

          This is the 1st line
          This is the 2nd line
          This is the 3rd line

      5  slc iterate(loop) over numbers 1-10 and output median (5.5)


    OBJECTIVE

      I am trying to make sql programmers expert programmers in
      9 languages using exactly the same sql queries.
      Use packages and procedures for analysis and sql for data wrangling and interfacing.

      Add very powerfull sql processing to the open source spss
      Posgresql has windows extensions.

      I hope to add repos with  drop downs to sql with windows extensions in many languages

     *language 1   open source spss
     *language 2   open source matlab
     *language 3   r
     *language 4   python
     *language 5   altair odbc sqlite
     *language 6   excel
      language 7   perl
      language 8   powershell

    /*                   _
    (_)_ __  _ __  _   _| |_
    | | `_ \| `_ \| | | | __|
    | | | | | |_) | |_| | |_
    |_|_| |_| .__/ \__,_|\__|
            |_|
    */

    * best first;
    %utlfkil(d:/sqlite/mysqlite.db);
    %utlfkil(d:/xls/class_export.xlsx);

    libname workx sas7bdat "d:/wpswrkx"; /*---  put in autoexec ---*/

    proc datasets lib=workx kill;
    run;

    libname sqlite odbc noprompt="driver=sqlite3 odbc driver; database=d:/sqlite/mysqlite.db;LoadExt=d:/dll/sqlean.dll;";
    libname xls excel "d:/xls/class_export.xlsx";

    proc datasets lib=sqlite;
     delete have;
    run;

    options validvarname=v7;
    data sqlite.have xls.have;
      informat
        line    $24.
        name    $8.
        sex     $1.
        age     8.
        height  8.
        ;
     input name sex age height line &;
    cards4;
    Alfred M 14 69 This is the 1st line
    Alice F 13 56.5 This is the 2nd line
    Barbara F 13 65.3 This is the 3rd line
    Carol F 14 62.8 This is the 4th line
    Henry M 14 63.5 This is the 5th line
    ;;;;
    run;quit;

    proc contents data=sqlite.have;
    run;

    proc print data=sqlite.have;
    run;

    proc sql;
     SELECT sql FROM sqlite.sqlite_master WHERE name='have'
    ;quit;

    libname sqlite clear;
    libname xls clear;

    /**************************************************************************************************************************/
    /* sqlite table have in database d:/sqlite/mysqlite.db                                                                    */
    /*                                                                                                                        */
    /* sqlite.have                                                                                                            */
    /*                                                                                                                        */
    /* line                        name        sex    age    height                                                           */
    /*                                                                                                                        */
    /* This is the 1st line        Alfred       M      14     69.0                                                            */
    /* This is the 2nd line        Alice        F      13     56.5                                                            */
    /* This is the 3rd line        Barbara      F      13     65.3                                                            */
    /* This is the 4th line        Carol        F      14     62.8                                                            */
    /* This is the 5th line        Henry        M      14     63.5                                                            */
    /*                                                                                                                        */
    /*                                                                                                                        */
    /* d:/xls/class_export.xlsx  sheet='have' (we will add she                                                                */
    /*                                                                                                                        */
    /* -----------------------+                                                                                               */
    /* | A1| fx    |DAYNUM    |                                                                                               */
    /* ---------------------------------------------------------------------------+                                           */
    /* [_] |    A     |    B    |    C    |   D     |          E                  |                                           */
    /* ---------------------------------------------------------------------------|                                           */
    /*  1  | NAME     |   SEX   |   AGE   | HEIGHT  |        LINE                 |                                           */
    /*  -- |----------+---------+---------+---------+-----------------------------|                                           */
    /*  2  |  Alfred  | M       | 14      | 69      | This is the 1st line        |                                           */
    /*  -- |----------+---------+---------+---------+-----------------------------|                                           */
    /*  3  |  Alice   | F       | 13      | 56.5    | This is the 2nd line        |                                           */
    /*  -- |----------+---------+---------+---------+-----------------------------|                                           */
    /*  4  |  Barbara | F       | 13      | 65.3    | This is the 3rd line        |                                           */
    /*  -- |----------+---------+---------+---------+-----------------------------|                                           */
    /*  5  |  Carol   | F       | 14      | 62.8    | This is the 4th line        |                                           */
    /*  -- |----------+---------+---------+---------+-----------------------------|                                           */
    /*  6  |  Henry   | M       | 14      | 63.5    | This is the 5th line        |                                           */
    /*  -- |----------+---------+---------+---------+-----------------------------|                                           */
    /* [HAVE]                                                                                                                 */
    /*                                                                                                                        */
    /* sqlite table have contents                                                                                             */
    /*                                                                                                                        */
    /* CREATE TABLE have( line VARCHAR(24), name VARCHAR(8), sex VARCHAR(1), age DOUBLE PRECISION, height DOUBLE PRECISION )  */
    /*                                                                                                                        */
    /* Altair SLC                                                                                                             */
    /*                                                                                                                        */
    /* The CONTENTS Procedure                                                                                                 */
    /*                                                                                                                        */
    /* Data Set Name           HAVE                                                                                           */
    /* Member Type             VIEW                                                                                           */
    /* Engine                                                                                                                 */
    /* Observations            .                                                                                              */
    /* Variables               5                                                                                              */
    /* Indexes                 0                                                                                              */
    /* Observation Length      49                                                                                             */
    /* Deleted Observations    0                                                                                              */
    /* Data Set Type                                                                                                          */
    /* Label                                                                                                                  */
    /* Compressed              NO                                                                                             */
    /* Sorted                  NO                                                                                             */
    /* Data Representation                                                                                                    */
    /* Encoding                wlatin1 Windows-1252 Western                                                                   */
    /*                                                                                                                        */
    /*       Alphabetic List of Variables and Attributes                                                                      */
    /*                                                                                                                        */
    /* Number    Variable    Type  Len   Pos    Format Informat                                                               */
    /* ________________________________________________________                                                               */
    /*      4    age         Num     8    33                                                                                  */
    /*      5    height      Num     8    41                                                                                  */
    /*      1    line        Char   24     0    $24.   $24.                                                                   */
    /*      2    name        Char    8    24    $8.    $8.                                                                    */
    /*      3    sex         Char    1    32    $1.    $1.                                                                    */
    /*                                                                                                                        */
    /**************************************************************************************************************************/

    /*                   _     _
    (_)_ __  _ __  _   _| |_  | | ___   __ _
    | | `_ \| `_ \| | | | __| | |/ _ \ / _` |
    | | | | | |_) | |_| | |_  | | (_) | (_| |
    |_|_| |_| .__/ \__,_|\__| |_|\___/ \__, |
            |_|                        |___/
    */

    1                                          Altair SLC        08:48 Thursday, March 19, 2026

    NOTE: Copyright 2002-2025 World Programming, an Altair Company
    NOTE: Altair SLC 2026 (05.26.01.00.000758)
          Licensed to Roger DeAngelis
    NOTE: This session is executing on the X64_WIN11PRO platform and is running in 64 bit mode

    NOTE: AUTOEXEC processing beginning; file is C:\wpsoto\autoexec.sas
    NOTE: AUTOEXEC source line
    1       +  ï»¿ods _all_ close;
               ^
    ERROR: Expected a statement keyword : found "?"


    NOTE: AUTOEXEC processing completed

    1         * best first;
    2         %utlfkil(d:/sqlite/mysqlite.db);
    3         %utlfkil(d:/xls/class_export.xlsx);
    4
    5         libname workx sas7bdat "d:/wpswrkx"; /*---  put in autoexec ---*/
    NOTE: Library workx assigned as follows:
          Engine:        SAS7BDAT
          Physical Name: d:\wpswrkx


    Altair SLC

    The DATASETS Procedure

             Directory

    Libref           WORKX
    Engine           SAS7BDAT
    Physical Name    d:\wpswrkx
    6
    7         proc datasets lib=workx kill;
    NOTE: No matching members in directory
    8         run;
    9
    10        libname sqlite odbc noprompt=XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX;
    NOTE: Library sqlite assigned as follows:
          Engine:        ODBC
          Physical Name:  (SQLite version 3.43.2)

    11        libname xls excel "d:/xls/class_export.xlsx";
    NOTE: Library xls assigned as follows:
          Engine:        OLEDB
          Physical Name: d:/xls/class_export.xlsx

    NOTE: Procedure datasets step took :
          real time : 0.551
          cpu time  : 0.609



    Altair SLC

    The DATASETS Procedure

          Directory

    Libref         SQLITE
    Engine         ODBC
    Data Source
    12
    13        proc datasets lib=sqlite;
    NOTE: No matching members in directory
    14         delete have;
    15        run;
    NOTE: SQLITE.HAVE (memtype="DATA") was not found, and has not been deleted
    16
    17        options validvarname=v7;
    NOTE: Procedure datasets step took :
          real time : 0.021
          cpu time  : 0.000


    18        data sqlite.have xls.have;
    19          informat
    20            line    $24.
    21            name    $8.
    22            sex     $1.
    23            age     8.
    24            height  8.
    25            ;
    26         input name sex age height line &;
    27        cards4;

    NOTE: Data set "SQLITE.have" has an unknown number of observation(s) and 5 variable(s)
    NOTE: Data set "XLS.have" has an unknown number of observation(s) and 5 variable(s)
    NOTE: The data step took :
          real time : 0.311
          cpu time  : 0.093


    28        Alfred M 14 69 This is the 1st line
    29        Alice F 13 56.5 This is the 2nd line
    30        Barbara F 13 65.3 This is the 3rd line
    31        Carol F 14 62.8 This is the 4th line
    32        Henry M 14 63.5 This is the 5th line
    33        ;;;;
    34        run;quit;
    35
    36        proc contents data=sqlite.have;
    37        run;
    NOTE: Procedure contents step took :
          real time : 0.063
          cpu time  : 0.031


    38
    39        proc print data=sqlite.have;
    40        run;
    NOTE: 5 observations were read from "SQLITE.have"
    NOTE: Procedure print step took :
          real time : 0.031
          cpu time  : 0.015


    41
    42        proc sql;
    43         SELECT sql FROM sqlite.sqlite_master WHERE name='have'
    44        ;quit;
    WARNING: truncating character column type to 1024 characters long, based on dbmax_text setting.
    WARNING: truncating character column name to 1024 characters long, based on dbmax_text setting.
    WARNING: truncating character column tbl_name to 1024 characters long, based on dbmax_text setting.
    WARNING: truncating character column sql to 1024 characters long, based on dbmax_text setting.
    WARNING: truncating character column sql to 1024 characters long, based on dbmax_text setting.
    NOTE: Procedure sql step took :
          real time : 0.015
          cpu time  : 0.000


    NOTE: Libref SQLITE has been deassigned.
    45
    46        libname sqlite clear;
    NOTE: Libref XLS has been deassigned.
    47        libname xls clear;
    ERROR: Error printed on page 1

    NOTE: Submitted statements took :
          real time : 1.108
          cpu time  : 0.843



    /*
     _ __  _ __ ___   ___ ___  ___ ___
    | `_ \| `__/ _ \ / __/ _ \/ __/ __|
    | |_) | | | (_) | (_|  __/\__ \__ \
    | .__/|_|  \___/ \___\___||___/___/
    |_|
    */

    %utlfkil(d:/txt/class_export.txt);

    options set=RHOME "C:\Progra~1\R\R-4.5.2\bin\r";

    proc r;
    export data=workx.have r=havex;
    submit;

    library(openxlsx)
    library(sqldf)
    options(sqldf.dll = "d:/dll/sqlean.dll")

    db_path="d:/sqlite/mysqlite.db"


    # function to add sheets
    add_sheet_to_existing_workbook <- function(xl_path, df, sheet_name) {
      # Load the existing workbook
      wb <- loadWorkbook(xl_path)

      # Add new sheet with data
      addWorksheet(wb, sheet_name)
      writeData(wb, sheet = sheet_name, x = df)

      # Save workbook (overwrites original)
      saveWorkbook(wb, xl_path, overwrite = TRUE)
    }

    # summarize have table
    avgs <- sqldf(c('
        select
           group_concat(name) as names
          ,sex
          ,round(avg(age),1)    as avg_age
          ,round(avg(height),1) as avg_height
        from
           have
        group
           by sex
        '), dbname = db_path);
    avgs

    # show that sqlite exists
    sqldf(c("SELECT sql FROM sqlite_master WHERE name='have'"), dbname = db_path)

    # create file d:/txt/class_export.txt uning lines column
    sqldf(c(
      "SELECT writefile(
         'd:/txt/class_export.txt',
         (SELECT group_concat(line, '\n') FROM have)
       )"
    ), dbname = db_path)

    # read back d:/txt/class_export.txt and create lines dataframe
    lines <- sqldf(c(
      "
      SELECT value AS txt
      FROM json_each('[\"' || replace(readfile('d:/txt/class_export.txt'), CHAR(10), '\",\"') || '\"]')
      "
    ), dbname = ":memory:")

    str(lines)
    print(lines)

    # loop through numbers 1-10 and compute the median
    median_value <- sqldf("
      WITH RECURSIVE nums(x) AS (
        SELECT 1
        UNION ALL
        SELECT x + 1 FROM nums WHERE x < 10
      ),
      med AS (
        SELECT
          x,
          ROW_NUMBER() OVER (ORDER BY x) AS r,
          COUNT(*) OVER () AS c
        FROM nums
      )
      SELECT AVG(x) AS median
      FROM med
      WHERE r IN ((c+1)/2, (c+2)/2)
    ")

    median_value
    print(paste("Median of 1-10:", median_value$median))

    xl_path="d:/xls/class_export.xlsx"

    #                             workbook    dataframe   sheet_name
    #                              -------    ---------   -----------
    add_sheet_to_existing_workbook(xl_path, median_value,"median_value")
    add_sheet_to_existing_workbook(xl_path, avgs, "avgs")
    add_sheet_to_existing_workbook(xl_path, lines, "lines")
    endsubmit;
    import r=avgs data=workx.avgs;
    import r=lines data=workx.lines;
    import r=median_value data=workx.median_value;
    run;

    proc print data=workx.avgs;
    run;

    proc print data=workx.lines;
    run;

    proc print data=workx.median_value;
    run;

    /**************************************************************************************************************************/
    /* d:/xls/class_export.xlsx  sheet='have' (we will add sheets)                                                            */
    /*                                                                                                                        */
    /* ORIGINAL SHEET HAVE EXISTS IN A CLOSED WORKBOOK (preserved style)                                                      */
    /* -------------------------------------                                                                                  */
    /*                                                                                                                        */
    /* -------------------------+                                                                                             */
    /* | A1| fx       | NAME    |                                                                                             */
    /* ---------------------------------------------------------------------------+                                           */
    /* [_] |    A     |    B    |    C    |   D     |          E                  |                                           */
    /* ---------------------------------------------------------------------------|                                           */
    /*  1  | NAME     |   SEX   |   AGE   | HEIGHT  |        LINE                 |                                           */
    /*  -- |----------+---------+---------+---------+-----------------------------|                                           */
    /*  2  |  Alfred  | M       | 14      | 69      | This is the 1st line        |                                           */
    /*  -- |----------+---------+---------+---------+-----------------------------|                                           */
    /*  3  |  Alice   | F       | 13      | 56.5    | This is the 2nd line        |                                           */
    /*  -- |----------+---------+---------+---------+-----------------------------|                                           */
    /*  4  |  Barbara | F       | 13      | 65.3    | This is the 3rd line        |                                           */
    /*  -- |----------+---------+---------+---------+-----------------------------|                                           */
    /*  5  |  Carol   | F       | 14      | 62.8    | This is the 4th line        |                                           */
    /*  -- |----------+---------+---------+---------+-----------------------------|                                           */
    /*  6  |  Henry   | M       | 14      | 63.5    | This is the 5th line        |                                           */
    /*  -- |----------+---------+---------+---------------------------------------|                                           */
    /* [HAVE]                                                                                                                 */
    /*                                                                                                                        */
    /*                                                                                                                        */
    /* CREATE TEXT FILE USING ONLY SQLITE WRITEFILE                                                                           */
    /* ---------------------------------------------                                                                          */
    /*                                                                                                                        */
    /*  d:/txt/class_export.txt                                                                                               */
    /*                                                                                                                        */
    /*  This is the 1st line                                                                                                  */
    /*  This is the 2nd line                                                                                                  */
    /*  This is the 3rd line                                                                                                  */
    /*  This is the 4th line                                                                                                  */
    /*  This is the 5th line                                                                                                  */
    /*                                                                                                                        */
    /*                                                                                                                        */
    /* ADD SHEET LINE USING ONLY SQLITE READFILE                                                                              */
    /* -----------------------------------------                                                                              */
    /*                                                                                                                        */
    /* -------------------------+                                                                                             */
    /* | A1| fx       | LINE    |                                                                                             */
    /* ----------------------------------+                                                                                    */
    /* [_] |           A                 |                                                                                    */
    /* ----------------------------------|                                                                                    */
    /*  1  |        LINE                 |                                                                                    */
    /*  -- |-----------------------------|                                                                                    */
    /*  2  | This is the 1st line        |                                                                                    */
    /*  -- |-----------------------------|                                                                                    */
    /*  3  | This is the 2nd line        |                                                                                    */
    /*  -- |-----------------------------|                                                                                    */
    /*  4  | This is the 3rd line        |                                                                                    */
    /*  -- |-----------------------------|                                                                                    */
    /*  5  | This is the 4th line        |                                                                                    */
    /*  -- |-----------------------------|                                                                                    */
    /*  6  | This is the 5th line        |                                                                                    */
    /*  -- |-----------------------------+                                                                                    */
    /* [LINE]                                                                                                                 */
    /*                                                                                                                        */
    /*                                                                                                                        */
    /* ADD SHEET AVGS                                                                                                         */
    /* ---------------                                                                                                        */
    /*                                                                                                                        */
    /* -----------------------------------+                                                                                   */
    /* | A1| fx                 |NAMES    |                                                                                   */
    /* ------------------------------------------------------------                                                           */
    /* [_] |    A               |    B    |    C    |      D      |                                                           */
    /* ------------------------------------------------------------                                                           */
    /*  1  | NAMES              |   SEX   |AVG_AGE  |AVG_ HEIGHT  |                                                           */
    /*  -- |--------------------+---------+---------+--------------                                                           */
    /*  2  | Alice,Barbara,Caro l       F |   13.3  |   61.5      |                                                           */
    /*  -- |----------+---------+---------+---------+--------------                                                           */
    /*  3  | Alfred,Henry       |       M |   14    |   66.3      |                                                           */
    /*  -- |--------------------+---------+---------+--------------                                                           */
    /* [AVGS]                                                                                                                 */
    /*                                                                                                                        */
    /**************************************************************************************************************************/

    /*
    | | ___   __ _
    | |/ _ \ / _` |
    | | (_) | (_| |
    |_|\___/ \__, |
             |___/
    */

    1                                          Altair SLC        10:06 Thursday, March 19, 2026

    NOTE: Copyright 2002-2025 World Programming, an Altair Company
    NOTE: Altair SLC 2026 (05.26.01.00.000758)
          Licensed to Roger DeAngelis
    NOTE: This session is executing on the X64_WIN11PRO platform and is running in 64 bit mode

    NOTE: AUTOEXEC processing beginning; file is C:\wpsoto\autoexec.sas
    NOTE: AUTOEXEC source line
    1       +  ï»¿ods _all_ close;
               ^
    ERROR: Expected a statement keyword : found "?"


    NOTE: AUTOEXEC processing completed

    1         %utlfkil(d:/txt/class_export.txt);
    2
    3         options set=RHOME "C:\Progra~1\R\R-4.5.2\bin\r";
    4
    5         proc r;
    NOTE: Using R version 4.5.2 (2025-10-31 ucrt) from C:\Program Files\R\R-4.5.2
    6         export data=workx.have r=havex;
                          ^
    ERROR: Data set "WORKX.have" not found
    7         submit;
    8
    9         library(openxlsx)
    10        library(sqldf)
    11        options(sqldf.dll = "d:/dll/sqlean.dll")
    12
    13        db_path="d:/sqlite/mysqlite.db"
    14
    15
    16        # function to add sheets
    17        add_sheet_to_existing_workbook <- function(xl_path, df, sheet_name) {
    18          # Load the existing workbook
    19          wb <- loadWorkbook(xl_path)
    20
    21          # Add new sheet with data
    22          addWorksheet(wb, sheet_name)
    23          writeData(wb, sheet = sheet_name, x = df)
    24
    25          # Save workbook (overwrites original)
    26          saveWorkbook(wb, xl_path, overwrite = TRUE)
    27        }
    28
    29        # summarize have table
    30        avgs <- sqldf(c('
    31            select
    32               group_concat(name) as names
    33              ,sex
    34              ,round(avg(age),1)    as avg_age
    35              ,round(avg(height),1) as avg_height
    36            from
    37               have
    38            group
    39               by sex
    40            '), dbname = db_path);
    41        avgs
    42
    43        # show that sqlite exists
    44        sqldf(c("SELECT sql FROM sqlite_master WHERE name='have'"), dbname = db_path)
    45
    46        # create file d:/txt/class_export.txt uning lines column
    47        sqldf(c(
    48          "SELECT writefile(
    49             'd:/txt/class_export.txt',
    50             (SELECT group_concat(line, '\n') FROM have)
    51           )"
    52        ), dbname = db_path)
    53
    54        # read back d:/txt/class_export.txt and create lines dataframe
    55        lines <- sqldf(c(
    56          "
    57          SELECT value AS txt
    58          FROM json_each('[\"' || replace(readfile('d:/txt/class_export.txt'), CHAR(10), '\",\"') || '\"]')
    59          "
    60        ), dbname = ":memory:")
    61
    62        str(lines)
    63        print(lines)
    64
    65        # loop through numbers 1-10 and compute the median
    66        median_value <- sqldf("
    67          WITH RECURSIVE nums(x) AS (
    68            SELECT 1
    69            UNION ALL
    70            SELECT x + 1 FROM nums WHERE x < 10
    71          ),
    72          med AS (
    73            SELECT
    74              x,
    75              ROW_NUMBER() OVER (ORDER BY x) AS r,
    76              COUNT(*) OVER () AS c
    77            FROM nums
    78          )
    79          SELECT AVG(x) AS median
    80          FROM med
    81          WHERE r IN ((c+1)/2, (c+2)/2)
    82        ")
    83
    84        median_value
    85        print(paste("Median of 1-10:", median_value$median))
    86
    87        xl_path="d:/xls/class_export.xlsx"
    88
    89        #                             workbook    dataframe   sheet_name
    90        #                              -------    ---------   -----------
    91        add_sheet_to_existing_workbook(xl_path, median_value,"median_value")
    92        add_sheet_to_existing_workbook(xl_path, avgs, "avgs")
    93        add_sheet_to_existing_workbook(xl_path, lines, "lines")
    94        endsubmit;

    NOTE: Submitting statements to R:

    >
    > library(openxlsx)
    > library(sqldf)
    Loading required package: gsubfn
    Loading required package: proto
    Loading required package: RSQLite
    > options(sqldf.dll = "d:/dll/sqlean.dll")
    >
    > db_path="d:/sqlite/mysqlite.db"
    >
    >
    > # function to add sheets
    > add_sheet_to_existing_workbook <- function(xl_path, df, sheet_name) {
    +   # Load the existing workbook
    +   wb <- loadWorkbook(xl_path)
    +
    +   # Add new sheet with data
    +   addWorksheet(wb, sheet_name)
    +   writeData(wb, sheet = sheet_name, x = df)
    +
    +   # Save workbook (overwrites original)
    +   saveWorkbook(wb, xl_path, overwrite = TRUE)
    + }
    >
    > # summarize have table
    > avgs <- sqldf(c('
    +     select
    +        group_concat(name) as names
    +       ,sex
    +       ,round(avg(age),1)    as avg_age
    +       ,round(avg(height),1) as avg_height
    +     from
    +        have
    +     group
    +        by sex
    +     '), dbname = db_path);
    > avgs
    >
    > # show that sqlite exists
    > sqldf(c("SELECT sql FROM sqlite_master WHERE name='have'"), dbname = db_path)
    >
    > # create file d:/txt/class_export.txt uning lines column
    > sqldf(c(
    +   "SELECT writefile(
    +      'd:/txt/class_export.txt',
    +      (SELECT group_concat(line, '\n') FROM have)
    +    )"
    + ), dbname = db_path)
    >
    > # read back d:/txt/class_export.txt and create lines dataframe
    > lines <- sqldf(c(
    +   "
    +   SELECT value AS txt
    +   FROM json_each('[\"' || replace(readfile('d:/txt/class_export.txt'), CHAR(10), '\",\"') || '\"]')
    +   "
    + ), dbname = ":memory:")
    >
    > str(lines)
    > print(lines)
    >
    > # loop through numbers 1-10 and compute the median
    > median_value <- sqldf("
    +   WITH RECURSIVE nums(x) AS (
    +     SELECT 1
    +     UNION ALL
    +     SELECT x + 1 FROM nums WHERE x < 10
    +   ),
    +   med AS (
    +     SELECT
    +       x,
    +       ROW_NUMBER() OVER (ORDER BY x) AS r,
    +       COUNT(*) OVER () AS c
    +     FROM nums
    +   )
    +   SELECT AVG(x) AS median
    +   FROM med
    +   WHERE r IN ((c+1)/2, (c+2)/2)
    + ")
    >
    > median_value
    > print(paste("Median of 1-10:", median_value$median))
    >
    > xl_path="d:/xls/class_export.xlsx"
    >
    > #                             workbook    dataframe   sheet_name
    > #                              -------    ---------   -----------
    > add_sheet_to_existing_workbook(xl_path, median_value,"median_value")
    > add_sheet_to_existing_workbook(xl_path, avgs, "avgs")
    > add_sheet_to_existing_workbook(xl_path, lines, "lines")
     
    NOTE: Processing of R statements complete

    95        import r=avgs data=workx.avgs;
    NOTE: Creating data set 'WORKX.avgs' from R data frame 'avgs'
    NOTE: Column names modified during import of 'avgs'
    NOTE: Data set "WORKX.avgs" has 2 observation(s) and 4 variable(s)

    96        import r=lines data=workx.lines;
    NOTE: Creating data set 'WORKX.lines' from R data frame 'lines'
    NOTE: Column names modified during import of 'lines'
    NOTE: Data set "WORKX.lines" has 5 observation(s) and 1 variable(s)

    97        import r=median_value data=workx.median_value;
    NOTE: Creating data set 'WORKX.median_value' from R data frame 'median_value'
    NOTE: Column names modified during import of 'median_value'
    NOTE: Data set "WORKX.median_value" has 1 observation(s) and 1 variable(s)

    NOTE: Step processing stopped because of errors detected
    98        run;
    NOTE: Procedure r step took :
          real time : 1.990
          cpu time  : 0.078


    99
    100       proc print data=workx.avgs;
    101       run;
    NOTE: 2 observations were read from "WORKX.avgs"
    NOTE: Procedure print step took :
          real time : 0.000
          cpu time  : 0.000


    102
    103       proc print data=workx.lines;
    104       run;
    NOTE: 5 observations were read from "WORKX.lines"
    NOTE: Procedure print step took :
          real time : 0.000
          cpu time  : 0.000


    105
    106       proc print data=workx.median_value;
    107       run;
    NOTE: 1 observations were read from "WORKX.median_value"
    NOTE: Procedure print step took :
          real time : 0.000
          cpu time  : 0.000


    ERROR: Error printed on page 1

    NOTE: Submitted statements took :
          real time : 2.116
          cpu time  : 0.187

    /*              _
      ___ _ __   __| |
     / _ \ `_ \ / _` |
    |  __/ | | | (_| |
     \___|_| |_|\__,_|
    */
