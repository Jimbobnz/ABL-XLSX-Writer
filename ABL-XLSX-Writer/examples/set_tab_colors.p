USING ABL-XLSX-Writer.ABL-XLSX-Writer.writer.*.

DEFINE VARIABLE objWorkbook     AS CLASS Workbook   NO-UNDO.
DEFINE VARIABLE objWorksheet    AS CLASS Worksheet  NO-UNDO.

objWorkbook  = NEW Workbook(INPUT THIS-PROCEDURE:FILE-NAME + '.xlsx').

/** Optional setting. **/
objWorkbook:set_properties('subject', "Set the Tab's colour").

objWorksheet = objWorkbook:add_worksheet().

objWorksheet = objWorkbook:add_worksheet().
objWorksheet:set_tab_color('red').

objWorksheet = objWorkbook:add_worksheet().
objWorksheet:set_tab_color('green').

objWorksheet = objWorkbook:add_worksheet().
objWorksheet:set_tab_color('blue').

MESSAGE objWorkbook:filename
    VIEW-AS ALERT-BOX INFO.
                                           
objWorkbook:close().





                    








