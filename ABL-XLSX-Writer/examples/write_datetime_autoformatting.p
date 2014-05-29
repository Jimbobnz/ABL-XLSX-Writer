
USING ABL-XLSX-Writer.ABL-XLSX-Writer.writer.*.

DEFINE VARIABLE objWorkbook     AS CLASS Workbook   NO-UNDO.
DEFINE VARIABLE objWorksheet    AS CLASS Worksheet  NO-UNDO.

DEFINE VARIABLE obj_Format      AS CLASS Format  NO-UNDO.


DEFINE VARIABLE param_DataTime  AS DATETIME         NO-UNDO. 

param_DataTime = NOW.

objWorkbook  = NEW Workbook(INPUT 'ABLExcelWriter.xlsx').

/** Optional setting. **/
objWorkbook:set_properties('title', 'ABL xslx Writer').
objWorkbook:set_properties('subject', 'My First Excel').
objWorkbook:set_properties('author', 'James ').
objWorkbook:set_properties('manager', 'James').
objWorkbook:set_properties('company', 'CMI').
objWorkbook:set_properties('keywords', 'First Excel').
objWorkbook:set_properties('comments', 'This is my very first Excel SpreadSheet').
objWorkbook:set_properties('category', 'Development').
objWorkbook:set_properties('status', 'Everything is Awsome').

objWorksheet = objWorkbook:add_worksheet('DateTime').

/** Set the the third column to be x width wide **/

obj_Format = objWorkbook:add_format().

obj_Format:set_bold().

objWorksheet:set_column(1,1,10).
objWorksheet:set_column(2,2,25).
objWorksheet:write(1,2, param_DataTime).
objWorksheet:write(1,1, 'Date Time:', obj_Format).      
    

objWorkbook:close().





                    









