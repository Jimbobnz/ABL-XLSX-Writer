USING ABL-XLSX-Writer.ABL-XLSX-Writer.writer.*.

DEFINE VARIABLE objWorkbook     AS CLASS Workbook   NO-UNDO.
DEFINE VARIABLE objWorksheet    AS CLASS Worksheet  NO-UNDO.
DEFINE VARIABLE ob_Format       AS CLASS FORMAT     NO-UNDO.
DEFINE VARIABLE ob_DateFormat   AS CLASS FORMAT     NO-UNDO.
DEFINE VARIABLE ob_TitleFormat  AS CLASS FORMAT     NO-UNDO.
DEFINE VARIABLE param_DataTime  AS DATETIME         NO-UNDO. 
DEFINE VARIABLE inLoop AS INTEGER     NO-UNDO.

param_DataTime = NOW.

objWorkbook  = NEW Workbook(INPUT THIS-PROCEDURE:FILE-NAME + '.xlsx').

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

objWorksheet = objWorkbook:add_worksheet('Header Footer').

objWorksheet:write(1,1,'TEST').
    
 objWorksheet:set_header('Header',?). 
 objWorksheet:set_footer('Footer',?). 
     

objWorkbook:close().





                    








