
USING ABL-XLSX-Writer.ABL-XLSX-Writer.writer.*.

DEFINE VARIABLE objWorkbook     AS CLASS Workbook   NO-UNDO.
DEFINE VARIABLE objWorksheet    AS CLASS Worksheet  NO-UNDO.

DEFINE VARIABLE obj_Format      AS CLASS Format  NO-UNDO.

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

objWorksheet = objWorkbook:add_worksheet('Formulas').

/** Set the the third column to be x width wide **/

obj_Format = objWorkbook:add_format('bold:true|font_colour:Grey|num_format:"0"|align:right').

objWorksheet:set_column(1,1,10).
objWorksheet:set_column(2,2,25).


objWorksheet:write(1,2, 1).      
objWorksheet:write(2,2, 2).      
objWorksheet:write(3,2, 3).      
objWorksheet:write(4,2, 4).      

objWorksheet:write(5,1, 'Total:', obj_Format).          
objWorksheet:write(5,2, '=SUM(B1:B4)', objWorkbook:add_format('font_colour:blue|bg_colour:orange|pattern:9')).          

objWorksheet:write(6,1, 'Average:', obj_Format).          
objWorksheet:write(6,2, '=AVERAGE(B1:B4)', obj_Format). 

objWorksheet:write(7,1, 'Minimum:', obj_Format).          
objWorksheet:write(7,2, '=MIN(B1:B4)', obj_Format). 

objWorksheet:write(8,1, 'Maximum:', obj_Format).          
objWorksheet:write(8,2, '=MAX(B1:B4)', obj_Format). 

objWorkbook:close().





                    









