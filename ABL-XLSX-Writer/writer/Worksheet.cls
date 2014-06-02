
/*     Copyright (C) 2014 James Bowen                                                                                                                       */
/*                                                                                                                                                          */
/*     Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"),   */
/*     to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense,   */
/*     and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:           */
/*                                                                                                                                                          */
/*     The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.                       */
/*                                                                                                                                                          */
/*     THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,  */
/*     FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER   */
/*     LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS */
/*     IN THE SOFTWARE.                                                                                                                                     */

USING ABL-XLSX-Writer.ABL-XLSX-Writer.writer.*.

&SCOPED-DEFINE da_EpocDate DATE(1,1,1900)
&SCOPED-DEFINE ximillisecondsInDay   0x5265C00 

CLASS ABL-XLSX-Writer.ABL-XLSX-Writer.writer.Worksheet INHERITS xmlwriter:
    
   DEFINE PUBLIC PROPERTY worksheetname             AS CHARACTER GET. SET.
   DEFINE PUBLIC PROPERTY optimization              AS LOGICAL INITIAL TRUE  GET. SET.
    
   DEFINE PUBLIC  PROPERTY SharedStringTable        AS CLASS SharedStringTable GET. SET. /** Object set by the workbook.cls **/
   DEFINE PRIVATE PROPERTY Utility                  AS CLASS Utility           GET. SET. 
   DEFINE PRIVATE PROPERTY Workbook                 AS CLASS Workbook          GET. SET. 
      
   DEFINE PRIVATE PROPERTY dim_rowmin               AS INTEGER INITIAL ? GET. SET.
   DEFINE PRIVATE PROPERTY dim_colmin               AS INTEGER INITIAL ? GET. SET. 
   DEFINE PRIVATE PROPERTY dim_rowmax               AS INTEGER GET. SET. 
   DEFINE PRIVATE PROPERTY dim_colmax               AS INTEGER GET. SET. 
   DEFINE PRIVATE PROPERTY span                     AS CHARACTER GET. SET.
   
    /** View Attributes **/
    DEFINE PUBLIC PROPERTY showGridLines            AS LOGICAL NO-UNDO  GET. PRIVATE SET. 
    DEFINE PUBLIC PROPERTY show_zeros               AS LOGICAL NO-UNDO  GET. PRIVATE SET.
    DEFINE PUBLIC PROPERTY is_right_to_left         AS LOGICAL NO-UNDO  GET. PRIVATE SET.
    DEFINE PUBLIC PROPERTY selected                 AS LOGICAL NO-UNDO  GET. SET.
    DEFINE PUBLIC PROPERTY outline_on               AS LOGICAL NO-UNDO  GET. PRIVATE SET.
    DEFINE PUBLIC PROPERTY page_view                AS LOGICAL NO-UNDO  GET. PRIVATE SET.    
    DEFINE PUBLIC PROPERTY zoom                     AS INTEGER NO-UNDO  GET. PRIVATE SET.    
    DEFINE PUBLIC PROPERTY zoom_scale_normal        AS LOGICAL NO-UNDO  GET. PRIVATE SET.    

    /** sheetFormatPr Attributes **/
    DEFINE PUBLIC PROPERTY default_row_height       AS INTEGER INITIAL 15     NO-UNDO GET. PRIVATE SET.
    DEFINE PUBLIC PROPERTY default_row_zeroed       AS LOGICAL                NO-UNDO GET. PRIVATE SET.
    DEFINE PUBLIC PROPERTY outline_row_level        AS INTEGER INITIAL 0      NO-UNDO GET. PRIVATE SET.
    DEFINE PUBLIC PROPERTY outline_col_level        AS INTEGER INITIAL 0      NO-UNDO GET. PRIVATE SET.
    DEFINE PUBLIC PROPERTY excel_version            AS INTEGER INITIAL 2007   NO-UNDO GET. PRIVATE SET.
    DEFINE PUBLIC PROPERTY tab_colour               AS CHARACTER              NO-UNDO GET. PRIVATE SET.
    
    /** pageSetup Attributes **/
    DEFINE PUBLIC PROPERTY paper_size                AS INTEGER   NO-UNDO                   GET. PRIVATE SET.
    
    DEFINE PUBLIC PROPERTY print_scale              AS INTEGER   NO-UNDO INITIAL 100        GET. SET.
    DEFINE PUBLIC PROPERTY orientation              AS CHARACTER NO-UNDO INITIAL 'portrait' GET. SET.
    DEFINE PUBLIC PROPERTY page_setup_changed       AS LOGICAL   NO-UNDO  GET. PRIVATE SET.    
    DEFINE PUBLIC PROPERTY print_options_changed    AS LOGICAL   NO-UNDO  GET. PRIVATE SET.  
    
    DEFINE PUBLIC PROPERTY margin_left              AS DECIMAL NO-UNDO GET. PRIVATE SET.  
    DEFINE PUBLIC PROPERTY margin_right             AS DECIMAL NO-UNDO GET. PRIVATE SET.  
    DEFINE PUBLIC PROPERTY margin_top               AS DECIMAL NO-UNDO GET. PRIVATE SET.
    DEFINE PUBLIC PROPERTY margin_bottom            AS DECIMAL NO-UNDO GET. PRIVATE SET.
    DEFINE PUBLIC PROPERTY margin_footer            AS DECIMAL NO-UNDO GET. PRIVATE SET.
    DEFINE PUBLIC PROPERTY margin_header            AS DECIMAL NO-UNDO GET. PRIVATE SET.

    DEFINE PUBLIC PROPERTY vcenter                  AS LOGICAL   NO-UNDO GET. PRIVATE SET.      
    DEFINE PUBLIC PROPERTY hcenter                  AS LOGICAL   NO-UNDO GET. PRIVATE SET.      
    
    DEFINE PUBLIC PROPERTY header_footer_changed    AS LOGICAL   NO-UNDO GET. PRIVATE SET.
    DEFINE PUBLIC PROPERTY print_headers            AS LOGICAL   NO-UNDO GET. PRIVATE SET.
    DEFINE PUBLIC PROPERTY header                   AS CHARACTER NO-UNDO GET. PRIVATE SET.
    DEFINE PUBLIC PROPERTY footer                   AS CHARACTER NO-UNDO GET. PRIVATE SET.
    
    DEFINE PUBLIC PROPERTY print_gridlines          AS LOGICAL   NO-UNDO GET. PRIVATE SET.
    DEFINE PUBLIC PROPERTY screen_gridlines         AS LOGICAL   NO-UNDO GET. PRIVATE SET.
    
    
    

    DEFINE PRIVATE TEMP-TABLE ttCols NO-UNDO SERIALIZE-NAME 'cols'
        FIELD min           AS INTEGER      XML-NODE-TYPE 'ATTRIBUTE' XML-NODE-NAME 'min'
        FIELD max           AS INTEGER      XML-NODE-TYPE 'ATTRIBUTE' XML-NODE-NAME 'max'
        FIELD width         AS DECIMAL      XML-NODE-TYPE 'ATTRIBUTE' XML-NODE-NAME 'width'
        FIELD objformat     AS CLASS Progress.Lang.OBJECT 
        INDEX idxCols is PRIMARY UNIQUE
            min
            max.

    DEFINE TEMP-TABLE ttMergeCells NO-UNDO SERIALIZE-NAME 'mergeCells'
        FIELD first_row     AS INTEGER      XML-NODE-TYPE 'HIDDEN'
        FIELD first_col     AS INTEGER      XML-NODE-TYPE 'HIDDEN'    
        FIELD last_row      AS INTEGER      XML-NODE-TYPE 'HIDDEN'
        FIELD last_col      AS INTEGER      XML-NODE-TYPE 'HIDDEN'
        FIELD ref           AS CHARACTER    XML-NODE-TYPE 'ATTRIBUTE' XML-NODE-NAME 'ref'.



    DEFINE PRIVATE TEMP-TABLE ttSheetDataRow NO-UNDO 
        FIELD CellRow       AS INTEGER
        FIELD xf_index      AS CHARACTER
        FIELD cell_format   AS CHARACTER
        FIELD height        AS INTEGER
        FIELD hidden        AS LOGICAL
        FIELD level         AS INTEGER
        FIELD collapsed     AS LOGICAL
        INDEX idxCellRow IS PRIMARY UNIQUE
            CellRow.

    DEFINE PRIVATE TEMP-TABLE ttSheetData NO-UNDO
      FIELD CellRow         AS INTEGER
      FIELD CellCol         AS INTEGER
      FIELD Cell_Value      AS CHARACTER
      FIELD DATATYPE        AS CHARACTER    INITIAL 'INTEGER'
      FIELD sst             AS INTEGER      INITIAL ?
      FIELD objformat       AS CLASS Progress.Lang.OBJECT 
      INDEX idxCellRowCol IS PRIMARY UNIQUE
            CellRow
            CellCol.
    
   CONSTRUCTOR Worksheet(INPUT obj_Workbook     AS CLASS Workbook,
                         INPUT pchWorkSheetName AS CHARACTER):
        THIS-OBJECT:Workbook      = obj_Workbook.  
        THIS-OBJECT:worksheetname = pchWorkSheetName.
        THIS-OBJECT:Utility       = NEW Utility().

        THIS-OBJECT:zoom                    = 100. /** Default the zoom to 100% **/
        
        THIS-OBJECT:optimization            = FALSE.

        THIS-OBJECT:default_row_height      = 15.
        THIS-OBJECT:default_row_zeroed      = FALSE.
        THIS-OBJECT:outline_row_level       = 0.
        THIS-OBJECT:outline_col_level       = 0.
        THIS-OBJECT:excel_version           = 2007. /** Default to Microsoft Excel 2007**/
        THIS-OBJECT:showGridLines           = TRUE.
        THIS-OBJECT:show_zeros              = TRUE.
        THIS-OBJECT:tab_colour              = ''.

        THIS-OBJECT:page_setup_changed      = FALSE.
        THIS-OBJECT:paper_size              = 0.
        THIS-OBJECT:orientation             = 'portrait'.

        THIS-OBJECT:print_options_changed   = FALSE.
        THIS-OBJECT:hcenter                 = FALSE.
        THIS-OBJECT:vcenter                 = FALSE.
        THIS-OBJECT:print_gridlines         = FALSE.
        THIS-OBJECT:screen_gridlines        = TRUE.
        THIS-OBJECT:print_headers           = FALSE.

        THIS-OBJECT:header_footer_changed   = FALSE.
        THIS-OBJECT:header                  = ''.
        THIS-OBJECT:footer                  = ''.

        THIS-OBJECT:margin_left             = .7.
        THIS-OBJECT:margin_right            = .7.
        THIS-OBJECT:margin_top              = .75.
        THIS-OBJECT:margin_bottom           = .75.
        THIS-OBJECT:margin_header           = .3.
        THIS-OBJECT:margin_footer           = .3.

   END CONSTRUCTOR.
    
   METHOD PRIVATE VOID _write (INPUT in_Row     AS INTEGER,
                               INPUT in_Col     AS INTEGER,
                               INPUT chVALUE    AS CHARACTER,
                               INPUT chDataType AS CHARACTER,
                               INPUT Obj_Format AS CLASS Format):

        DEFINE VARIABLE inShared_String_Index AS INTEGER     NO-UNDO.

        FIND ttSheetDataRow
            WHERE ttSheetDataRow.CellRow EQ in_Row
            NO-ERROR.

        IF NOT AVAILABLE ttSheetDataRow THEN
            CREATE ttSheetDataRow.

        ASSIGN 
            ttSheetDataRow.CellRow = in_Row.
        
        IF VALID-OBJECT(Obj_Format) THEN
        DO:
            IF chDataType EQ 'DATETIME' AND Obj_Format:num_format EQ '0' THEN
                Obj_Format:set_num_Format(INPUT 'd mmmm yyyy hh:mm:ss').
            
            IF chDataType EQ 'DECIMAL' AND Obj_Format:num_format EQ '0' THEN
                Obj_Format:set_num_Format(INPUT '0.00'). /** Option Number 2**/
            
            IF chDataType EQ 'INTEGER' THEN
                Obj_Format:set_num_Format(INPUT '0').    /** Option Number 1**/

            IF chDataType EQ 'CHARACTER' AND Obj_Format:num_format EQ '0' THEN
                Obj_Format:set_num_Format(INPUT 'General').    /** Option Number 1**/

        END.            

        FIND ttSheetData
            WHERE ttSheetData.CellRow EQ in_Row
            AND   ttSheetData.CellCol EQ in_Col
            NO-ERROR.

        IF NOT AVAILABLE ttSheetData THEN
        DO:
            CREATE ttSheetData.

            ASSIGN 
                ttSheetData.CellRow   = in_Row   
                ttSheetData.CellCol   = in_Col.

            IF VALID-OBJECT(Obj_Format) THEN
                ttSheetData.objformat = Obj_Format.
        END.

        IF dim_rowmin EQ ? THEN
            THIS-OBJECT:dim_rowmin = in_Row.

        IF dim_colmin EQ ? THEN
            THIS-OBJECT:dim_colmin = in_Col.

        THIS-OBJECT:dim_rowmin  = MINIMUM(THIS-OBJECT:dim_rowmin, in_Row).
        THIS-OBJECT:dim_colmin  = MINIMUM(THIS-OBJECT:dim_colmin, in_Col).
        THIS-OBJECT:dim_rowmax  = MAXIMUM(THIS-OBJECT:dim_rowmax, in_Row).
        THIS-OBJECT:dim_colmax  = MAXIMUM(THIS-OBJECT:dim_colmax, in_Col).

        /** For string optimization all string/character values go into a common string table.**/
        IF chDataType EQ 'CHARACTER' THEN
            chVALUE = STRING( THIS-OBJECT:SharedStringTable:_get_shared_string_index( chVALUE ) ).

        ASSIGN
            ttSheetData.Cell_Value  = chVALUE
            ttSheetData.DATATYPE    = chDataType.
   END METHOD.

    METHOD PUBLIC VOID set_tab_color(INPUT pch_tab_colour AS CHARACTER):
        THIS-OBJECT:tab_colour = THIS-OBJECT:Utility:xl_colour(INPUT pch_tab_colour).
    END METHOD.

    METHOD PUBLIC VOID protect (INPUT pch_password AS CHARACTER):
    END METHOD.

/* ########################################################################### */
/* #                                                                           */
/* # Public API. Page Setup methods.                                           */
/* #                                                                           */
/* ########################################################################### */

    METHOD PUBLIC VOID set_landscape():
        THIS-OBJECT:orientation        = 'landscape'.
        THIS-OBJECT:page_setup_changed = TRUE.
    END METHOD.

    METHOD PUBLIC VOID set_portrait():
        THIS-OBJECT:orientation        = 'portrait'.
        THIS-OBJECT:page_setup_changed = TRUE.
    END METHOD.

    METHOD PUBLIC VOID set_page_view():
        THIS-OBJECT:page_view = TRUE.
    END METHOD.
    
    METHOD PUBLIC VOID set_paper(INPUT pc_paper_size AS CHARACTER):
         THIS-OBJECT:paper_size         = paper_size.
         THIS-OBJECT:page_setup_changed = TRUE.
    END METHOD.
    
    METHOD PUBLIC VOID center_horizontally():
        THIS-OBJECT:hcenter               = TRUE.
        THIS-OBJECT:print_options_changed = TRUE.
    END METHOD.

    METHOD PUBLIC VOID center_vertically():
        THIS-OBJECT:vcenter               = TRUE.
        THIS-OBJECT:print_options_changed = TRUE.
    END METHOD.

    METHOD PUBLIC VOID set_margins (INPUT in_left   AS DECIMAL,
                                    INPUT in_Right  AS DECIMAL, 
                                    INPUT in_top    AS DECIMAL,
                                    INPUT in_Bottom AS DECIMAL):

        /** Assign to the defaults.**/
        ASSIGN  in_left   = 0.75 WHEN (NOT LOGICAL(in_left)).
        ASSIGN  in_Right  = 0.75 WHEN (NOT LOGICAL(in_Right)).
        ASSIGN  in_top    = 0.75 WHEN (NOT LOGICAL(in_top)).
        ASSIGN  in_bottom = 0.75 WHEN (NOT LOGICAL(in_bottom)).
        
        THIS-OBJECT:margin_left   = in_left.
        THIS-OBJECT:margin_right  = in_Right.
        THIS-OBJECT:margin_top    = in_top.    
        THIS-OBJECT:margin_bottom = in_Bottom.

    END METHOD.

    METHOD PUBLIC VOID set_header(INPUT pch_Headers AS CHARACTER,
                                  INPUT in_Margin   AS DECIMAL):

        IF LENGTH(pch_Headers) GT 255 THEN
        DO:
            MESSAGE 
                'Header string must be less than 255 characters'
                VIEW-AS ALERT-BOX WARNING.
            RETURN.
        END.
            

        ASSIGN in_Margin = 0.3 WHEN (in_Margin EQ ?).

        THIS-OBJECT:header                = pch_Headers.
        THIS-OBJECT:margin_header         = in_Margin.
        THIS-OBJECT:header_footer_changed = TRUE.
    END METHOD.

    METHOD PUBLIC VOID set_footer(INPUT pch_footer AS CHARACTER,
                                  INPUT in_Margin  AS DECIMAL):

        IF LENGTH(pch_footer) GT 255 THEN
        DO:
            MESSAGE 
                'Footer string must be less than 255 characters'
                VIEW-AS ALERT-BOX WARNING.
            RETURN.
        END.
            
        ASSIGN in_Margin = 0.3 WHEN (in_Margin EQ ?).

        THIS-OBJECT:footer                = pch_footer.
        THIS-OBJECT:margin_footer         = in_Margin.
        THIS-OBJECT:header_footer_changed = TRUE.
    END METHOD.
    

    METHOD PUBLIC VOID hide_gridlines(INPUT pin_Option AS INTEGER):
        
        CASE pin_Option:
            WHEN 0 THEN
            DO:
                THIS-OBJECT:print_gridlines       = TRUE.
                THIS-OBJECT:screen_gridlines      = TRUE.
                THIS-OBJECT:print_options_changed = TRUE.
            END.
            WHEN 1 THEN
            DO:
                THIS-OBJECT:print_gridlines       = FALSE.
                THIS-OBJECT:screen_gridlines      = TRUE.
            END.
            OTHERWISE
            DO:
                THIS-OBJECT:print_gridlines       = FALSE.
                THIS-OBJECT:screen_gridlines      = FALSE.
            END.
        END CASE.

    END METHOD.


        
    
   /* Merge a range of cells. */
   METHOD PUBLIC VOID merge_range(INPUT in_first_row  AS INTEGER,
                                  INPUT in_first_col  AS INTEGER,
                                  INPUT in_last_row   AS INTEGER,
                                  INPUT in_last_col   AS INTEGER):

        FIND FIRST ttMergeCells
            WHERE (ttMergeCells.first_row LE in_first_row AND
                   ttMergeCells.last_row  GE in_first_row AND  
                   ttMergeCells.first_col LE in_first_col AND
                   ttMergeCells.last_col  GE in_first_col) OR 
                  (ttMergeCells.first_row LE in_last_row AND
                   ttMergeCells.last_row  GE in_last_row AND  
                   ttMergeCells.first_col LE in_last_col AND
                   ttMergeCells.last_col  GE in_last_col)
                NO-ERROR.

        IF AVAILABLE ttMergeCells THEN
        DO:
            MESSAGE 'Over-lapping range detected.'.
            RETURN.
        END.

        CREATE ttMergeCells.

        ASSIGN
            ttMergeCells.first_row = in_first_row 
            ttMergeCells.first_col = in_first_col 
            ttMergeCells.last_row  = in_last_row  
            ttMergeCells.last_col  = in_last_col.  

        /** Excel Cell Notation 'A1:B2' **/
        ASSIGN ttMergeCells.ref = SUBSTITUTE('&1:&2',
                                             THIS-OBJECT:Utility:xl_rowcol_to_cell(INPUT in_first_row,
                                                                                   INPUT in_first_col),
                                             THIS-OBJECT:Utility:xl_rowcol_to_cell(INPUT in_last_row,
                                                                                   INPUT in_last_col)
                                             ).

   END METHOD.


   METHOD PUBLIC VOID set_column(INPUT in_First_Col  AS INTEGER,
                                 INPUT in_Last_Col   AS INTEGER,
                                 INPUT de_Witdh      AS DECIMAL):

        DEFINE VARIABLE ob_Col_Format AS CLASS FORMAT.
        ob_Col_Format = THIS-OBJECT:Workbook:add_Format().

        THIS-OBJECT:set_column(INPUT in_First_Col,  
                               INPUT in_Last_Col,  
                               INPUT de_Witdh,
                               INPUT ob_Col_Format).

   END METHOD.

   METHOD PUBLIC VOID set_column(INPUT in_First_Col  AS INTEGER,
                                 INPUT in_Last_Col   AS INTEGER,
                                 INPUT de_Witdh      AS DECIMAL,
                                 INPUT ob_Col_Format AS CLASS FORMAT):

        DEFINE VARIABLE First_Col AS INTEGER   NO-UNDO.
        DEFINE VARIABLE Last_Col  AS INTEGER   NO-UNDO.

/*        DEFINE PRIVATE TEMP-TABLE ttCols NO-UNDO SERIALIZE-NAME 'cols' */
/*         FIELD min       AS INTEGER XML-NODE-NAME 'min'                */
/*         FIELD max       AS INTEGER XML-NODE-NAME 'max'                */
/*         FIELD width     AS DECIMAL XML-NODE-NAME 'width'              */
/*         FIELD objformat AS CLASS Progress.Lang.OBJECT                 */
/*         INDEX idxCols is PRIMARY UNIQUE                               */
/*             min                                                       */
/*             max.                                                      */

        ASSIGN
            First_Col = MINIMUM(1,in_First_Col) 
            Last_Col  = MAXIMUM(1,in_Last_Col).  

        /** Transpose values arround **/
        IF First_Col GT Last_Col THEN
            ASSIGN
                in_First_Col = Last_Col 
                in_Last_Col  = First_Col. 

        FIND ttCols 
            WHERE ttCols.min EQ in_First_Col
            AND   ttCols.max EQ in_Last_Col
            NO-ERROR.

        IF NOT AVAILABLE ttCols THEN
        DO:
            CREATE ttCols.
            ASSIGN 
                ttCols.min = in_First_Col
                ttCols.max = in_Last_Col.
        END.

        ASSIGN
            ttCols.width = de_Witdh.

        IF VALID-OBJECT(ob_Col_Format) THEN
            ASSIGN
                ttCols.objformat = ob_Col_Format.

   END METHOD.

    /** Write LOGICAL without format object **/
    METHOD PUBLIC VOID write (INPUT in_Row  AS INTEGER,
                              INPUT in_Col  AS INTEGER,
                              INPUT lgVALUE AS LOGICAL):

        IF (in_Row GT 1048576) OR
           (in_Col GT 16384) THEN
            RETURN.

        DEFINE VARIABLE obj_Format AS CLASS FORMAT.
        
        obj_Format = THIS-OBJECT:Workbook:add_Format().

        THIS-OBJECT:write(INPUT in_Row, 
                          INPUT in_Col,
                          INPUT lgVALUE, /* 1=TRUE, 0,FALSE*/
                          INPUT obj_Format).
    END METHOD.

    /** Write LOGICAL with format object **/
    METHOD PUBLIC VOID write (INPUT in_Row     AS INTEGER,
                              INPUT in_Col     AS INTEGER,
                              INPUT lgVALUE    AS LOGICAL,
                              INPUT obj_Format AS CLASS FORMAT):

        IF (in_Row GT 1048576) OR
           (in_Col GT 16384) THEN
            RETURN.

        THIS-OBJECT:_write(INPUT in_Row, 
                          INPUT in_Col,
                          INPUT STRING(lgVALUE,'1/0'), /* 1=TRUE, 0,FALSE*/
                          INPUT 'LOGICAL',
                          INPUT obj_Format).
    END METHOD.

    /** Write CHARACTER without format object **/
    METHOD PUBLIC VOID write (INPUT in_Row  AS INTEGER,
                              INPUT in_Col  AS INTEGER,
                              INPUT chVALUE AS CHARACTER):

        IF (in_Row GT 1048576) OR
            (in_Col GT 16384) THEN
            RETURN.
        
        DEFINE VARIABLE obj_Format AS CLASS FORMAT.
        
        obj_Format = THIS-OBJECT:Workbook:add_Format().
        
        THIS-OBJECT:write(INPUT in_Row,
                          INPUT in_Col,
                          INPUT chVALUE,
                          INPUT obj_Format).
    END METHOD.
    
    /** Write CHARACTER without format object **/
    METHOD PUBLIC VOID write (INPUT in_Row      AS INTEGER,
                              INPUT in_Col      AS INTEGER,
                              INPUT chVALUE     AS CHARACTER,
                              INPUT obj_Format  AS CLASS FORMAT):

        IF (in_Row GT 1048576) OR
           (in_Col GT 16384) THEN
            RETURN.
        
        /** Test the expression of the value to see if its a formula. **/
        IF chVALUE BEGINS '=' THEN
        DO:

            chVALUE = REPLACE(chVALUE,'=','').

            THIS-OBJECT:_write(INPUT in_Row, 
                               INPUT in_Col,
                               INPUT chVALUE,
                               INPUT 'FORMULA':U,
                               INPUT obj_Format).
        END.
        ELSE
            THIS-OBJECT:_write(INPUT in_Row, 
                               INPUT in_Col,
                               INPUT chVALUE,
                               INPUT 'CHARACTER',
                               INPUT obj_Format).

    END METHOD.
    
    /** Write integer without format object **/
    METHOD PUBLIC VOID write (INPUT in_Row      AS INTEGER,
                              INPUT in_Col      AS INTEGER,
                              INPUT inVALUE     AS INTEGER):


        IF (in_Row GT 1048576) OR
           (in_Col GT 16384) THEN
            RETURN.

        DEFINE VARIABLE obj_Format AS CLASS FORMAT.

        obj_Format = THIS-OBJECT:Workbook:add_Format().
    
        THIS-OBJECT:write(INPUT in_Row, 
                          INPUT in_Col,
                          INPUT inVALUE,
                          INPUT obj_Format).
    END METHOD.

    /** Write integer with format object **/
    METHOD PUBLIC VOID write (INPUT in_Row      AS INTEGER,
                              INPUT in_Col      AS INTEGER,
                              INPUT inVALUE     AS INTEGER,
                              INPUT obj_Format  AS CLASS FORMAT):

        IF (in_Row GT 1048576) OR
           (in_Col GT 16384) THEN
            RETURN.
    
        THIS-OBJECT:_write(INPUT in_Row, 
                           INPUT in_Col,
                           INPUT STRING(inVALUE),
                           INPUT 'INTEGER',
                           INPUT obj_Format).
    END METHOD.

    /** Write decimal without format object **/
    METHOD PUBLIC VOID write (INPUT in_Row   AS INTEGER,
                              INPUT in_Col   AS INTEGER,
                              INPUT deVALUE  AS DECIMAL):

        IF (in_Row GT 1048576) OR
           (in_Col GT 16384) THEN
            RETURN.

        DEFINE VARIABLE obj_Format AS CLASS FORMAT.
        obj_Format = THIS-OBJECT:Workbook:add_Format().

        THIS-OBJECT:_write (INPUT in_Row, 
                            INPUT in_Col,
                            INPUT STRING(deVALUE),
                            INPUT 'DECIMAL',
                            INPUT obj_Format).

    END METHOD.

    /** Write decimal with format object **/
    METHOD PUBLIC VOID write (INPUT in_Row     AS INTEGER,
                              INPUT in_Col     AS INTEGER,
                              INPUT deVALUE    AS DECIMAL,
                              INPUT obj_Format AS CLASS FORMAT):

        IF (in_Row GT 1048576) OR
           (in_Col GT 16384) THEN
            RETURN.

        THIS-OBJECT:_write(INPUT in_Row, 
                           INPUT in_Col,
                           INPUT STRING(deVALUE),
                           INPUT 'DECIMAL',
                           INPUT obj_Format).

    END METHOD.

    /** Write DATE without format object **/
    METHOD PUBLIC VOID write (INPUT in_Row   AS INTEGER,
                              INPUT in_Col   AS INTEGER,
                              INPUT daVALUE  AS DATE):

        DEFINE VARIABLE in_Interval  AS INT64    NO-UNDO.
        DEFINE VARIABLE obj_Format   AS CLASS FORMAT.

        obj_Format = THIS-OBJECT:Workbook:add_Format().

        THIS-OBJECT:write(INPUT in_Row, 
                           INPUT in_Col,
                           INPUT in_Interval,
                           INPUT obj_Format).
    END METHOD.


    /** Write DATE with format object **/
    METHOD PUBLIC VOID write (INPUT in_Row      AS INTEGER,
                              INPUT in_Col      AS INTEGER,
                              INPUT daVALUE     AS DATE,
                              INPUT obj_Format   AS CLASS FORMAT):

        DEFINE VARIABLE in_Interval AS INT64     NO-UNDO.
        
        /** + 2 Include the Epoc date and the date passed. **/
        in_Interval = INTERVAL(daVALUE, {&da_EpocDate},'days') + 2.

        THIS-OBJECT:_write(INPUT in_Row, 
                           INPUT in_Col,
                           INPUT STRING(in_Interval),
                           INPUT 'DATE',
                           INPUT obj_Format).
    END METHOD.
    
    /** Write DATETIME without format object **/
    METHOD PUBLIC VOID write (INPUT in_Row   AS INTEGER,
                              INPUT in_Col   AS INTEGER,
                              INPUT dt_Value AS DATETIME):
        
        IF dt_Value EQ ? THEN
            RETURN.

        DEFINE VARIABLE obj_Format AS CLASS FORMAT.

        obj_Format = THIS-OBJECT:Workbook:add_Format().

        THIS-OBJECT:write(INPUT in_Row,
                          INPUT in_Col,
                          INPUT dt_Value,
                          INPUT obj_Format).
    END METHOD.

    /** Write DATETIME with format object **/
    METHOD PUBLIC VOID write (INPUT in_Row    AS INTEGER,
                              INPUT in_Col    AS INTEGER,
                              INPUT dt_Value  AS DATETIME,
                              INPUT obj_Format AS CLASS FORMAT):

        DEFINE VARIABLE in_Interval AS INT64     NO-UNDO.
        DEFINE VARIABLE deTime_Dec  AS DECIMAL   NO-UNDO.
        DEFINE VARIABLE ch_DateTime AS CHARACTER NO-UNDO.

        IF dt_Value EQ ? THEN
            RETURN.

        /**NOTE: + 2 Include the Epoc date and the date passed. **/
        in_Interval = INTERVAL(DATE(dt_Value), {&da_EpocDate}, 'days') + 2.

/*         deTime_Dec  = MTIME(dt_Value) / (60 * 60 * 24 * 1000). */
        deTime_Dec  = MTIME(dt_Value) / {&ximillisecondsInDay}.
        ch_DateTime = STRING(in_Interval) + STRING(deTime_Dec).

        THIS-OBJECT:_write(INPUT in_Row, 
                           INPUT in_Col,
                           INPUT ch_DateTime,
                           INPUT 'DATETIME',
                           INPUT obj_Format).


    END METHOD.

    METHOD PUBLIC VOID write (INPUT in_Row   AS INTEGER,
                              INPUT in_Col   AS INTEGER,
                              INPUT tz_Value AS DATETIME-TZ):
            
        DEFINE VARIABLE param_DateTime AS DATETIME        NO-UNDO.

        param_DateTime = DATETIME( DATE(tz_Value), MTIME(tz_Value) ).

        /** Call the method above...**/
        THIS-OBJECT:write (INPUT in_Row,
                           INPUT in_Col,
                           INPUT param_DateTime).

    END.

    METHOD PUBLIC VOID write (INPUT in_Row      AS INTEGER,
                              INPUT in_Col      AS INTEGER,
                              INPUT tz_Value    AS DATETIME-TZ,
                              INPUT obj_Format   AS CLASS FORMAT):

        DEFINE VARIABLE param_DateTime AS DATETIME        NO-UNDO.
    
        param_DateTime = DATETIME( DATE(tz_Value), MTIME(tz_Value) ).
    
        /** Call the method above...**/
        THIS-OBJECT:write (INPUT in_Row,
                           INPUT in_Col,
                           INPUT param_DateTime,
                           INPUT obj_Format).
    
    END.



    METHOD PUBLIC VOID write_temp_table(INPUT in_row        AS INTEGER, 
                                        INPUT in_col        AS INTEGER,
                                        INPUT hn_temptable  AS HANDLE).

            /**
            PUBLIC METHOD TO take the handle a converting it into excell Fields. 
            */

            IF NOT VALID-HANDLE(hn_temptable) OR
               NOT hn_temptable:type EQ 'TEMP-TABLE' THEN
                RETURN.

            DEFINE VARIABLE hn_Buffer       AS HANDLE      NO-UNDO.
            DEFINE VARIABLE hn_Query        AS HANDLE      NO-UNDO.
            DEFINE VARIABLE in_Num_fields   AS INTEGER     NO-UNDO.
            DEFINE VARIABLE in_Field_Loop   AS INTEGER     NO-UNDO.
            DEFINE VARIABLE ch_field_value  AS CHARACTER   NO-UNDO.

            hn_Buffer = hn_temptable:DEFAULT-BUFFER-HANDLE.
            in_Num_fields = hn_Buffer:NUM-FIELDS.

            DO in_Field_Loop = 1 TO in_Num_fields :

                ch_field_value = hn_Buffer:BUFFER-FIELD( in_Field_Loop ):COLUMN-LABEL.

                IF LOGICAL(hn_Buffer:BUFFER-FIELD( in_Field_Loop ):EXTENT) THEN
                    ch_field_value = SUBSTITUTE('&1 [&2]',
                                                ch_field_value,
                                                hn_Buffer:BUFFER-FIELD( in_Field_Loop ):EXTENT).

                THIS-OBJECT:write(INPUT in_row,
                                  INPUT in_col + (in_Field_Loop - 1),
                                  INPUT ch_field_value ).

            END.

            CREATE QUERY hn_Query.
            hn_Query:ADD-BUFFER(hn_Buffer).
            hn_Query:QUERY-PREPARE( SUBSTITUTE('FOR EACH &1', hn_temptable:NAME ) ).
            hn_Query:QUERY-OPEN().

            IF hn_Query:IS-OPEN THEN
            DO:

                hn_Query:GET-FIRST() NO-ERROR.
                REPEAT WHILE NOT hn_Query:QUERY-OFF-END:

                    IF NOT hn_Buffer:AVAILABLE THEN 
                        NEXT.

                    PROCESS EVENTS.

                    /** Next Row..**/
                    in_row = in_row + 1.
                    DO in_Field_Loop = 1 TO in_Num_fields:

                        /** Can't handle arrays yet..**/

                        IF LOGICAL(hn_Buffer:BUFFER-FIELD( in_Field_Loop ):EXTENT) THEN
                            NEXT.

                        CASE hn_Buffer:BUFFER-FIELD( in_Field_Loop ):DATA-TYPE:
                            WHEN 'INTEGER' THEN
                                THIS-OBJECT:write(INPUT in_row, 
                                                  INPUT in_col + (in_Field_Loop - 1), 
                                                  INPUT INTEGER(hn_Buffer:BUFFER-FIELD( in_Field_Loop ):BUFFER-VALUE())).
                            WHEN 'DECIMAL' THEN
                                THIS-OBJECT:write(INPUT in_row, 
                                                  INPUT in_col + (in_Field_Loop - 1), 
                                                  INPUT DECIMAL(hn_Buffer:BUFFER-FIELD( in_Field_Loop ):BUFFER-VALUE())).
                            OTHERWISE
                                    THIS-OBJECT:write(INPUT in_row,
                                                      INPUT in_col + (in_Field_Loop - 1), 
                                                      INPUT TRIM(STRING(hn_Buffer:BUFFER-FIELD( in_Field_Loop ):STRING-VALUE()))).
                        END CASE.
                    END.

                    hn_Query:GET-NEXT() NO-ERROR.
                END.

                hn_Query:QUERY-CLOSE().
            END.

            DELETE OBJECT hn_Query.
    END METHOD.

   METHOD PUBLIC VOID _assemble_xml_file():

        /** Write the workbook element. **/
        THIS-OBJECT:_write_worksheet().

        /** Close the file. **/
        THIS-OBJECT:_xml_close().

   END METHOD.

/*########################################################################### */
/*#                                                                           */
/*# XML methods.                                                              */
/*#                                                                           */
/*########################################################################### */


   METHOD PRIVATE VOID _write_worksheet().

    hnSAXWriter:START-ELEMENT('worksheet'). 
    hnSAXWriter:DECLARE-NAMESPACE('http://schemas.openxmlformats.org/spreadsheetml/2006/main').
    hnSAXWriter:DECLARE-NAMESPACE('http://schemas.openxmlformats.org/officeDocument/2006/relationships','r').
    
    /** Add some extra attributes for Excel 2010. Mainly for sparklines.**/
    IF THIS-OBJECT:excel_version EQ 2010 THEN
    DO:
        hnSAXWriter:DECLARE-NAMESPACE('http://schemas.openxmlformats.org/markup-compatibility/2006','mc').
        hnSAXWriter:DECLARE-NAMESPACE('http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac','x14ac').
        hnSAXWriter:INSERT-ATTRIBUTE('mc:Ignorable','x14ac').
    END.
    
    
    THIS-OBJECT:_write_sheet_pr().
    THIS-OBJECT:_write_dimension().
    THIS-OBJECT:_write_sheet_views().
    THIS-OBJECT:_write_sheet_format_pr().
    THIS-OBJECT:_write_cols().

    IF THIS-OBJECT:optimization THEN
        THIS-OBJECT:_write_optimized_sheet_data().        
    ELSE
        THIS-OBJECT:_write_sheet_data().
        
    THIS-OBJECT:_write_sheet_protection().
    THIS-OBJECT:_write_auto_filter().
    THIS-OBJECT:_write_merge_cells().
/*     # Write the conditional formats. */
    THIS-OBJECT:_write_conditional_formats().
/*     # Write the dataValidations element. */
    THIS-OBJECT:_write_data_validations().
/*     # Write the hyperlink element. */
    THIS-OBJECT:_write_hyperlinks().
/*     # Write the printOptions element. */
    THIS-OBJECT:_write_print_options().
/*     # Write the worksheet page_margins. */
    THIS-OBJECT:_write_page_margins().
/*     # Write the worksheet page setup. */
    THIS-OBJECT:_write_page_setup().
/*     # Write the headerFooter element. */
    THIS-OBJECT:_write_header_footer().
/*     # Write the rowBreaks element. */
    THIS-OBJECT:_write_row_breaks().
/*     # Write the colBreaks element. */
    THIS-OBJECT:_write_col_breaks().
/*     # Write the drawing element. */
    THIS-OBJECT:_write_drawings().
/*     # Write the legacyDrawing element. */
    THIS-OBJECT:_write_legacy_drawing().
/*     # Write the tableParts element. */
    THIS-OBJECT:_write_table_parts().
/*     # Write the extLst and sparklines. */
    THIS-OBJECT:_write_ext_sparklines().

    hnSAXWriter:END-ELEMENT('worksheet').   
   END METHOD.

   METHOD PRIVATE VOID _write_sheet_data():

        IF NOT TEMP-TABLE ttSheetDataRow:HAS-RECORDS THEN
           hnSAXWriter:WRITE-EMPTY-ELEMENT('sheetData').      
        ELSE
        DO:
            hnSAXWriter:START-ELEMENT('sheetData').  

            THIS-OBJECT:_write_rows().

            hnSAXWriter:END-ELEMENT('sheetData').  
        END.
   END METHOD.

   METHOD PRIVATE VOID _write_rows():

        DEFINE VARIABLE obj_Format        AS CLASS FORMAT.
        DEFINE VARIABLE obj_Format_Instant AS CLASS Format.

        THIS-OBJECT:_calculate_spans().

        FOR EACH ttSheetDataRow:

            hnSAXWriter:START-ELEMENT('row').
            hnSAXWriter:INSERT-ATTRIBUTE('r', STRING(ttSheetDataRow.CellRow) ).
            hnSAXWriter:INSERT-ATTRIBUTE('spans', THIS-OBJECT:span ).

            FOR EACH ttSheetData
                WHERE ttSheetData.CellRow EQ ttSheetDataRow.CellRow
                    BY ttSheetData.CellCol:
            
                hnSAXWriter:START-ELEMENT('c').
                hnSAXWriter:INSERT-ATTRIBUTE('r', THIS-OBJECT:Utility:xl_rowcol_to_cell(ttSheetData.CellRow, ttSheetData.CellCol)).

                obj_Format_Instant = CAST(ttSheetData.ObjFormat , format ).
                
                IF VALID-OBJECT(obj_Format_Instant) THEN
                DO:

                    DEFINE VARIABLE in_xf_Index AS INTEGER     NO-UNDO.

                    in_xf_Index = obj_Format_Instant:_get_xf_index().

                    IF obj_Format_Instant:_get_xf_index() NE 0 THEN 
                        hnSAXWriter:INSERT-ATTRIBUTE('s', STRING( in_xf_Index ) ).
                END.
                    
                /** http://officeopenxml.com/SScontentOverview.php **/

/*                 Valid value Description   */
/*                 b           Boolean       */
/*                 n           Number        */
/*                 e           Error         */
/*                 s           Shared String */
/*                 str         String        */
/*                 inlineStr   Inline String */

                CASE ttSheetData.DataType:
                    WHEN 'CHARACTER' THEN
                        hnSAXWriter:INSERT-ATTRIBUTE('t', 's'). /** s = string lookup table..**/
                    WHEN 'INTEGER' OR 
                    WHEN 'DECIMAL' THEN
                        hnSAXWriter:INSERT-ATTRIBUTE('t', 'n'). /** n = number**/
                    WHEN 'LOGICAL' THEN
                        hnSAXWriter:INSERT-ATTRIBUTE('t', 'b'). /** b = boolian **/
                    WHEN 'FORMULA' THEN
                        hnSAXWriter:INSERT-ATTRIBUTE('t', 'str'). /** str = inline string **/
                    OTHERWISE
                        hnSAXWriter:INSERT-ATTRIBUTE('t', 'n'). /** n = number **/
                END CASE.

                IF ttSheetData.DatATYPE EQ 'FORMULA' THEN
                    hnSAXWriter:WRITE-DATA-ELEMENT('f', ttSheetData.Cell_Value).
                ELSE
                    IF ttSheetData.Cell_Value NE '' THEN
                        hnSAXWriter:WRITE-DATA-ELEMENT('v', ttSheetData.Cell_Value).

                hnSAXWriter:END-ELEMENT('c').
       
            END.

            hnSAXWriter:END-ELEMENT('row').

        END.

   END METHOD.

   METHOD PRIVATE VOID _calculate_spans():

        /* */

        DEFINE VARIABLE inMin_Col AS INTEGER INITIAL ? NO-UNDO.
        DEFINE VARIABLE inMaxCol AS INTEGER INITIAL ? NO-UNDO.


        FOR EACH ttSheetDataRow
            WHERE ttSheetDataRow.CellRow LE 16,
            EACH ttSheetData
                WHERE ttSheetData.CellRow EQ ttSheetDataRow.CellRow
                BY ttSheetData.CellCol:
                    
                IF inMin_Col EQ ? THEN
                    inMin_Col = ttSheetData.CellCol.
                
                IF inMaxCol EQ ? THEN
                    inMaxCol = ttSheetData.CellCol.

                INMin_Col = MINIMUM(INMin_Col, ttSheetData.CellCol).
                inMaxCol  = MAXIMUM(inMaxCol, ttSheetData.CellCol).

            
        END.

        THIS-OBJECT:span = SUBSTITUTE("&1:&2",
                                      inMin_Col,
                                      inMaxCol).

   END METHOD.

   METHOD PRIVATE VOID _write_optimized_sheet_data():
   END METHOD.

    METHOD PRIVATE VOID _write_sheet_pr():
        /** If  tab colour has not been defined then return. **/
        IF THIS-OBJECT:tab_colour EQ '' THEN
            RETURN.

        hnSAXWriter:START-ELEMENT('sheetPr').

        THIS-OBJECT:_write_tab_color().

        hnSAXWriter:END-ELEMENT('sheetPr').
    END METHOD.

    METHOD PRIVATE VOID _write_tab_color():
        hnSAXWriter:WRITE-EMPTY-ELEMENT('tabColor' ).
        hnSAXWriter:INSERT-ATTRIBUTE('rgb', THIS-OBJECT:tab_colour).
    END METHOD.




   METHOD PRIVATE VOID _write_dimension():

        DEFINE VARIABLE ref AS CHARACTER   NO-UNDO.

        IF THIS-OBJECT:dim_rowmin EQ ? AND
           THIS-OBJECT:dim_colmin EQ ? THEN
            ref = "A1".
        ELSE 
        IF THIS-OBJECT:dim_rowmin EQ 0 AND THIS-OBJECT:dim_colmin GT 0 THEN
        DO:


            IF THIS-OBJECT:dim_colmin EQ THIS-OBJECT:dim_colmax THEN
                ref = THIS-OBJECT:Utility:xl_rowcol_to_cell(0, THIS-OBJECT:dim_colmin).
            ELSE
                ref = SUBSTITUTE('&1:&2',
                                 THIS-OBJECT:Utility:xl_rowcol_to_cell(0, THIS-OBJECT:dim_colmin),
                                 THIS-OBJECT:Utility:xl_rowcol_to_cell(0, THIS-OBJECT:dim_colmax)
                                 ).
            
        END.
        ELSE 
        IF (THIS-OBJECT:dim_rowmin EQ THIS-OBJECT:dim_rowmax AND
            THIS-OBJECT:dim_colmin EQ THIS-OBJECT:dim_colmax) THEN
            ref = THIS-OBJECT:Utility:xl_rowcol_to_cell(THIS-OBJECT:dim_rowmin, THIS-OBJECT:dim_colmin).
        ELSE
            ref = SUBSTITUTE('&1:&2',
                             THIS-OBJECT:Utility:xl_rowcol_to_cell(THIS-OBJECT:dim_rowmin, THIS-OBJECT:dim_colmin),
                             THIS-OBJECT:Utility:xl_rowcol_to_cell(THIS-OBJECT:dim_rowmax, THIS-OBJECT:dim_colmax)
                             ).

        hnSAXWriter:WRITE-EMPTY-ELEMENT('dimension').       
        hnSAXWriter:INSERT-ATTRIBUTE('ref', ref ).       
   END METHOD.

   METHOD PRIVATE VOID _write_sheet_views():

        hnSAXWriter:START-ELEMENT('sheetViews'). 

        THIS-OBJECT:_write_sheet_view().

        hnSAXWriter:END-ELEMENT('sheetViews'). 
   END METHOD.

    METHOD PRIVATE VOID _write_sheet_view():
        
        /*self.panes or len(self.selections):*/

        hnSAXWriter:WRITE-EMPTY-ELEMENT('sheetView':U).       

        IF NOT THIS-OBJECT:showGridLines THEN
            hnSAXWriter:INSERT-ATTRIBUTE('showGridLines':U, '0' ). 

        IF NOT THIS-OBJECT:show_zeros THEN
            hnSAXWriter:INSERT-ATTRIBUTE('showZeros':U, '0' ). 

        IF THIS-OBJECT:is_right_to_left THEN
            hnSAXWriter:INSERT-ATTRIBUTE('rightToLeft':U, '1' ). 

        IF THIS-OBJECT:selected THEN
            hnSAXWriter:INSERT-ATTRIBUTE('tabSelected':U, '1' ).  

        IF NOT THIS-OBJECT:outline_on THEN
            hnSAXWriter:INSERT-ATTRIBUTE('showOutlineSymbols':U, '0' ).  

        IF THIS-OBJECT:page_view THEN
            hnSAXWriter:INSERT-ATTRIBUTE('view', 'pageLayout' ).  

        IF THIS-OBJECT:zoom NE 100 THEN
        DO:
            IF NOT THIS-OBJECT:page_view THEN
            DO:
                hnSAXWriter:INSERT-ATTRIBUTE('zoomScale':U, STRING(THIS-OBJECT:zoom) ). 

                IF THIS-OBJECT:zoom_scale_normal THEN
                    hnSAXWriter:INSERT-ATTRIBUTE('zoomScaleNormal':U, STRING(THIS-OBJECT:zoom) ). 
            END.
        END.

        hnSAXWriter:INSERT-ATTRIBUTE('workbookViewId':U, '0').

   END METHOD.

   METHOD PRIVATE VOID _write_sheet_format_pr():

        hnSAXWriter:WRITE-EMPTY-ELEMENT('sheetFormatPr').   
        
        hnSAXWriter:INSERT-ATTRIBUTE('defaultRowHeight':U, STRING(THIS-OBJECT:default_row_height) ).

        IF THIS-OBJECT:default_row_height NE 15 THEN
            hnSAXWriter:INSERT-ATTRIBUTE('customHeight':U, '1' ).

        IF THIS-OBJECT:default_row_zeroed THEN
            hnSAXWriter:INSERT-ATTRIBUTE('zeroHeight':U, '1' ).
        
        IF LOGICAL(THIS-OBJECT:outline_row_level) THEN
            hnSAXWriter:INSERT-ATTRIBUTE('outlineLevelRow':U, STRING( THIS-OBJECT:outline_row_level) ).

        IF LOGICAL(THIS-OBJECT:outline_col_level) THEN
            hnSAXWriter:INSERT-ATTRIBUTE('outlineLevelCol':U, STRING( THIS-OBJECT:outline_col_level) ).

        IF THIS-OBJECT:excel_version EQ 2010 THEN
            hnSAXWriter:INSERT-ATTRIBUTE('x14ac:dyDescent':U, '0.25' ).

   END METHOD.
   
    METHOD PRIVATE VOID _write_cols():
        
        IF NOT TEMP-TABLE ttCols:HAS-RECORDS THEN
            RETURN.

        hnSAXWriter:START-ELEMENT('cols').   
        THIS-OBJECT:_write_col_info(). 
        hnSAXWriter:END-ELEMENT('cols').   
        
    END METHOD.

    METHOD PRIVATE VOID _write_col_info():

        DEFINE VARIABLE custom_width AS LOGICAL INITIAL FALSE NO-UNDO.

        FOR EACH ttCols:
            
            IF ttCols.width EQ ? THEN
                ASSIGN
                    ttCols.width = 8.43      /** 8.43 is the default..**/
                    custom_width = FALSE.    
            
            hnSAXWriter:WRITE-EMPTY-ELEMENT('col').   
        
            hnSAXWriter:INSERT-ATTRIBUTE('min':U, STRING( ttCols.min ) ).
            hnSAXWriter:INSERT-ATTRIBUTE('max':U, STRING( ttCols.max ) ).
            hnSAXWriter:INSERT-ATTRIBUTE('width':U, STRING( ttCols.width ) ).
        
            IF ttCols.width NE 8.43 THEN
                hnSAXWriter:INSERT-ATTRIBUTE('customWidth':U, '1' ).
        
        END.

    END METHOD.

   METHOD PRIVATE VOID _write_sheet_protection():
   END METHOD.

   METHOD PRIVATE VOID _write_auto_filter():
   END METHOD.

    METHOD PRIVATE VOID _write_merge_cells():
        
        DEFINE VARIABLE hnttMergeCells  AS HANDLE      NO-UNDO.
        DEFINE VARIABLE hnMergeCell     AS HANDLE      NO-UNDO.
        DEFINE VARIABLE hnQuery         AS HANDLE      NO-UNDO.

        hnttMergeCells = TEMP-TABLE ttMergeCells:HANDLE.

        IF NOT hnttMergeCells:HAS-RECORDS THEN
            RETURN.

        hnMergeCell = hnttMergeCells:DEFAULT-BUFFER-HANDLE.

        CREATE QUERY hnQuery.
        hnQuery:ADD-BUFFER(hnMergeCell).
        hnQuery:QUERY-PREPARE('PRESELECT EACH ttMergeCells').
        hnQuery:QUERY-OPEN().

        IF hnQuery:IS-OPEN THEN
        DO:
            hnQuery:GET-FIRST() NO-ERROR.
    
            hnSAXWriter:START-ELEMENT('mergeCells').   
            hnSAXWriter:INSERT-ATTRIBUTE('count':U, STRING( hnQuery:NUM-RESULTS ) ).
    
            REPEAT WHILE NOT hnQuery:QUERY-OFF-END:
    
                THIS-OBJECT:_write_merge_cell(INPUT hnMergeCell).
                hnQuery:GET-NEXT() NO-ERROR.
            END.
    
            hnSAXWriter:END-ELEMENT('mergeCells').   
    
            hnQuery:QUERY-CLOSE().
        END.

        DELETE OBJECT hnQuery.

   END METHOD.

    METHOD PRIVATE VOID _write_merge_cell(INPUT hnBuffer AS HANDLE):

        IF NOT VALID-HANDLE(hnBuffer)   OR
           NOT hnBuffer:TYPE EQ 'BUFFER' THEN
           RETURN.

        hnSAXWriter:WRITE-EMPTY-ELEMENT('mergeCell').  
        hnSAXWriter:INSERT-ATTRIBUTE('ref', STRING(hnBuffer::ref) ).   

   END METHOD.

   METHOD PRIVATE VOID _write_conditional_formats():
   END METHOD.

   METHOD PRIVATE VOID _write_data_validations():
   END METHOD.

   METHOD PRIVATE VOID _write_hyperlinks():
   END METHOD.

   METHOD PRIVATE VOID _write_print_options():

        IF NOT THIS-OBJECT:print_options_changed THEN
            RETURN.

        hnSAXWriter:WRITE-EMPTY-ELEMENT('printOptions'). 

/*         # Set horizontal centering. */
        IF THIS-OBJECT:hcenter THEN
            hnSAXWriter:INSERT-ATTRIBUTE('horizontalCentered', '1').

/*         # Set vertical centering. */
        IF THIS-OBJECT:vcenter THEN
            hnSAXWriter:INSERT-ATTRIBUTE('verticalCentered', '1').

/*         # Enable row and column headers. */
        IF THIS-OBJECT:print_headers THEN
            hnSAXWriter:INSERT-ATTRIBUTE('headings',' 1').

/*         # Set printed gridlines. */
        IF THIS-OBJECT:print_gridlines THEN
            hnSAXWriter:INSERT-ATTRIBUTE('gridLines', '1').
        
   END METHOD.

   METHOD PRIVATE VOID _write_page_margins():

        hnSAXWriter:WRITE-EMPTY-ELEMENT('pageMargins').   
        hnSAXWriter:INSERT-ATTRIBUTE('left',    TRIM(STRING(THIS-OBJECT:margin_left, '>>9.99'))).   
        hnSAXWriter:INSERT-ATTRIBUTE('right',   TRIM(STRING(THIS-OBJECT:margin_right, '>>9.99'))).   
        hnSAXWriter:INSERT-ATTRIBUTE('top',     TRIM(STRING(THIS-OBJECT:margin_top, '>>9.99'))).  
        hnSAXWriter:INSERT-ATTRIBUTE('bottom',  TRIM(STRING(THIS-OBJECT:margin_bottom, '>>9.99'))).   

        hnSAXWriter:INSERT-ATTRIBUTE('header',  TRIM(STRING(THIS-OBJECT:margin_header, '>>9.99'))).   
        hnSAXWriter:INSERT-ATTRIBUTE('footer',  TRIM(STRING(THIS-OBJECT:margin_footer, '>>9.99'))).   

   END METHOD.

    METHOD PRIVATE VOID _write_page_setup():

        IF NOT THIS-OBJECT:page_setup_changed THEN
            RETURN.

        hnSAXWriter:WRITE-EMPTY-ELEMENT('pageSetup').   

        IF LOGICAL(THIS-OBJECT:paper_size)  THEN
            hnSAXWriter:INSERT-ATTRIBUTE('paper_size', STRING(THIS-OBJECT:paper_size) ).   

        IF THIS-OBJECT:print_scale NE 100 THEN /** 100% **/
            hnSAXWriter:INSERT-ATTRIBUTE('scale', STRING( THIS-OBJECT:print_scale ) ).   
        
        hnSAXWriter:INSERT-ATTRIBUTE('orientation', THIS-OBJECT:orientation ).   
        
   END METHOD.

   METHOD PRIVATE VOID _write_header_footer():

        IF NOT THIS-OBJECT:header_footer_changed THEN
            RETURN.

        hnSAXWriter:START-ELEMENT('headerFooter'). 

        IF THIS-OBJECT:header NE '' THEN
            THIS-OBJECT:_write_odd_header().

        IF THIS-OBJECT:footer NE '' THEN
            THIS-OBJECT:_write_odd_footer().

        hnSAXWriter:END-ELEMENT('headerFooter'). 

   END METHOD.

    METHOD PRIVATE VOID _write_odd_header():
        hnSAXWriter:WRITE-DATA-ELEMENT('oddHeader', THIS-OBJECT:header ).   
    END METHOD.

    METHOD PRIVATE VOID _write_odd_footer():
        hnSAXWriter:WRITE-DATA-ELEMENT('oddFooter', THIS-OBJECT:footer ).   
    END METHOD.


   METHOD PRIVATE VOID _write_row_breaks():
   END METHOD.

   METHOD PRIVATE VOID _write_col_breaks():
   END METHOD.

   METHOD PRIVATE VOID _write_drawings():
   END METHOD.

   METHOD PRIVATE VOID _write_legacy_drawing():
   END METHOD.

   METHOD PRIVATE VOID _write_table_parts():
   END METHOD.

   METHOD PRIVATE VOID _write_ext_sparklines():
   END METHOD.

END CLASS.
