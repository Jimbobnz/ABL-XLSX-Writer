
USING ABL-XLSX-Writer.ABL-XLSX-Writer.writer.*.

DEFINE VARIABLE objWorkbook     AS CLASS Workbook   NO-UNDO.
DEFINE VARIABLE objWorksheet    AS CLASS Worksheet  NO-UNDO.

DEFINE VARIABLE obj_Format      AS CLASS Format  NO-UNDO.

objWorkbook  = NEW Workbook(INPUT THIS-PROCEDURE:FILE-NAME + '.xlsx').

/** Optional setting. **/
objWorkbook:set_properties('title', 'ABL xslx Writer').
objWorkbook:set_properties('subject', 'Patterns').
objWorkbook:set_properties('author', 'James ').
objWorkbook:set_properties('manager', 'James').
objWorkbook:set_properties('company', 'CMI').
objWorkbook:set_properties('keywords', 'First Excel').
objWorkbook:set_properties('comments', 'This is my very first Excel SpreadSheet').
objWorkbook:set_properties('category', 'Development').
objWorkbook:set_properties('status', 'Everything is Awsome').                                           

objWorksheet = objWorkbook:add_worksheet('Patterns').

/** Set the the third column to be x width wide **/

objWorksheet:set_column(2,2,25).

objWorksheet:write(1,2, 1, objWorkbook:add_format('align:centre|font_colour:white|bg_colour:orange|pattern:1|rotation:90')).      
objWorksheet:write(2,2, 2, objWorkbook:add_format('align:centre|font_colour:white|bg_colour:orange|pattern:2|rotation:80')).      
objWorksheet:write(3,2, 3, objWorkbook:add_format('align:centre|font_colour:white|bg_colour:orange|pattern:3|rotation:70')).      
objWorksheet:write(4,2, 4, objWorkbook:add_format('align:centre|font_colour:white|bg_colour:orange|pattern:4|rotation:60')).      
objWorksheet:write(5,2, 5, objWorkbook:add_format('align:centre|font_colour:white|bg_colour:orange|pattern:5|rotation:50')).      
objWorksheet:write(6,2, 6, objWorkbook:add_format('align:centre|font_colour:white|bg_colour:orange|pattern:6|rotation:40')).      
objWorksheet:write(7,2, 7, objWorkbook:add_format('align:centre|font_colour:white|bg_colour:orange|pattern:7|rotation:30')).      
objWorksheet:write(8,2, 8, objWorkbook:add_format('align:centre|font_colour:white|bg_colour:orange|pattern:8|rotation:20')).      
objWorksheet:write(9,2, 9, objWorkbook:add_format('align:centre|font_colour:white|bg_colour:orange|pattern:9|rotation:10')).      
objWorksheet:write(10,2, 10, objWorkbook:add_format('align:centre|font_colour:white|bg_colour:orange|pattern:10|rotation:0')).      
objWorksheet:write(11,2, 11, objWorkbook:add_format('align:centre|font_colour:white|bg_colour:orange|pattern:11|rotation:-10')).      
objWorksheet:write(12,2, 12, objWorkbook:add_format('align:centre|font_colour:white|bg_colour:orange|pattern:12|rotation:-20')).      
objWorksheet:write(13,2, 13, objWorkbook:add_format('align:centre|font_colour:white|bg_colour:orange|pattern:13|rotation:-30')).      
objWorksheet:write(14,2, 14, objWorkbook:add_format('align:centre|font_colour:white|bg_colour:orange|pattern:14|rotation:-40')).      
objWorksheet:write(15,2, 15, objWorkbook:add_format('align:centre|font_colour:white|bg_colour:orange|pattern:15|rotation:-50')).      
objWorksheet:write(16,2, 16, objWorkbook:add_format('align:centre|font_colour:white|bg_colour:orange|pattern:16|rotation:-60')).      
objWorksheet:write(17,2, 17, objWorkbook:add_format('align:centre|font_colour:white|bg_colour:orange|pattern:17|rotation:-70')).      
objWorksheet:write(18,2, 18, objWorkbook:add_format('align:centre|font_colour:white|bg_colour:orange|pattern:18|rotation:-80')).      

objWorkbook:close().
