USING ABL-XLSX-Writer\ABL-XLSX-Writer.writer*.

DEFINE VARIABLE objWorkbook     AS CLASS Workbook   NO-UNDO.
DEFINE VARIABLE objWorksheet    AS CLASS Worksheet  NO-UNDO.
DEFINE VARIABLE ob_Format       AS CLASS FORMAT     NO-UNDO.
DEFINE VARIABLE ob_DateFormat   AS CLASS FORMAT     NO-UNDO.
DEFINE VARIABLE ob_TitleFormat  AS CLASS FORMAT     NO-UNDO.
DEFINE VARIABLE param_DataTime  AS DATETIME         NO-UNDO. 
DEFINE VARIABLE inLoop AS INTEGER     NO-UNDO.

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

objWorksheet = objWorkbook:add_worksheet('Clients').

    inLoop = 3.

    /** Set the the third column to be x width wide **/
    
    objWorksheet:set_column(2,2,21.86).
    objWorksheet:set_column(3,3,31).

    ob_TitleFormat = objWorkbook:add_Format().
    ob_TitleFormat:set_font_size(12).
    ob_TitleFormat:set_bold().
  

    objWorksheet:write(1,1, 'THis is my FIRST xslx Writer', ob_TitleFormat ).  

    objWorksheet:merge_range(1,1,1,3).
    
    ob_Format = objWorkbook:add_Format().
    ob_Format:set_bold().
    ob_Format:set_italic().
    ob_Format:set_font_colour('red').
    ob_Format:set_align('center').
    ob_Format:set_align('vcenter').
    ob_Format:set_rotation(45).

    
    ob_DateFormat = objWorkbook:add_Format().
    ob_DateFormat:set_num_Format(INPUT 'd mmmm yyyy').
    ob_DateFormat:set_align('vcenter').

    FOR EACH Product   NO-LOCK:
        
         objWorksheet:write(inLoop,1, Product.ProductOrder  ).  
         objWorksheet:write(inLoop,2, Product.ProductName, ob_Format). 
         objWorksheet:write(inLoop,3, Product.Created, ob_DateFormat). 
         inLoop = inLoop + 1.

    END.

objWorkbook:close().





                    








