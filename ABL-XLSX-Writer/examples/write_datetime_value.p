USING ABL-XLSX-Writer.ABL-XLSX-Writer.writer.*.

DEFINE VARIABLE objWorkbook          AS CLASS Workbook   NO-UNDO.
DEFINE VARIABLE objWorksheet         AS CLASS Worksheet  NO-UNDO.
DEFINE VARIABLE obj_DateTime_Format  AS CLASS FORMAT  NO-UNDO.
DEFINE VARIABLE obj_Text_Format      AS CLASS FORMAT  NO-UNDO.

DEFINE VARIABLE da_date_time_now AS DATETIME  NO-UNDO.

da_date_time_now = NOW.

/** Create a new Workbook object and supplied it the of this program and appended .xlsx.**/
objWorkbook  = NEW Workbook(INPUT THIS-PROCEDURE:FILE-NAME + '.xlsx').

/** Optional. Setting the workbook's property for the subject name. **/
objWorkbook:set_properties('subject', "Formatted Date Time").

/** Create a new worksheet and give it a name of 'Date Time'. **/
objWorksheet        = objWorkbook:add_worksheet('Date Time').

/** Set the worksheets tab colour to red. Could also use a HEX coded colour format '#FF5520'.**/
objWorksheet:set_tab_color('red').

/** Create an new format object **/
  
obj_DateTime_Format = objWorkbook:add_format().
obj_Text_Format     = objWorkbook:add_format().  



/** Turn on bold formatting **/
obj_Text_Format:set_bold().
/** Set the format's object alignment to 'right' align.**/
obj_Text_Format:set_align('right').
obj_DateTime_Format:set_align('center').
/** Set the format's number format to day month year hours minutes and seconds **/
obj_DateTime_Format:set_num_format('d mmm yyyy h:mm:ss').

/** Write a character string to row 1, column 1 ('A1') and apply it's formatting..**/
objWorksheet:write(1,1,'Formatted Date:', obj_Text_Format).

/** Write the DATETIME variable to row 1, column 2 ('A2') and apply it's formatting..**/
objWorksheet:write(1,2, da_date_time_now, obj_DateTime_Format).

/** Set the column with for both column 1 & 2 to size 30.**/
objWorksheet:set_column(1,2,30).

/** Close the workbook and package it up into an .xlsx document.**/
objWorkbook:close().

/** Show the filename of the xlsx document that been created using the filename property.**/
MESSAGE objWorkbook:filename VIEW-AS ALERT-BOX INFO.
