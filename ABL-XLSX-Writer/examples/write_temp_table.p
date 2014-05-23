USING ABL-XLSX-Writer.ABL-XLSX-Writer.writer.*.

DEFINE VARIABLE objWorkbook          AS CLASS Workbook   NO-UNDO.
DEFINE VARIABLE objWorksheet         AS CLASS Worksheet  NO-UNDO.
DEFINE VARIABLE obj_DateTime_Format  AS CLASS FORMAT  NO-UNDO.
DEFINE VARIABLE obj_Text_Format      AS CLASS FORMAT  NO-UNDO.

DEFINE VARIABLE inCounter AS INTEGER     NO-UNDO.

DEFINE VARIABLE hn_TempTable AS HANDLE      NO-UNDO.

DEFINE TEMP-TABLE ttCallCentreOutcome LIKE CallCentreOutcome.

FOR EACH CallCentreOutcome:

    CREATE ttCallCentreOutcome.
    BUFFER-COPY CallCentreOutcome TO ttCallCentreOutcome.

    inCounter = inCounter + 1.

END.

MESSAGE
    inCounter
    view-as alert-box info.

hn_TempTable = TEMP-TABLE ttCallCentreOutcome:HANDLE.


/** Create a new Workbook object and supplied it the of this program and appended .xlsx.**/
objWorkbook  = NEW Workbook(INPUT THIS-PROCEDURE:FILE-NAME + '.xlsx').

/** Optional. Setting the workbook's property for the subject name. **/
objWorkbook:set_properties('subject', "Formatted Date Time").

/** Create a new worksheet and give it a name of 'Date Time'. **/
objWorksheet        = objWorkbook:add_worksheet('TEMP-TABLE').

/** Set the worksheets tab colour to red. Could also use a HEX coded colour format '#FF5520'.**/
objWorksheet:set_tab_color('blue').

objWorksheet:write_temp_table(2,2,INPUT hn_TempTable).

/** Close the workbook and package it up into an .xlsx document.**/
objWorkbook:close().

/** Show the filename of the xlsx document that been created using the filename property.**/
MESSAGE objWorkbook:filename VIEW-AS ALERT-BOX INFO.
