
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
USING Progress.Json.ObjectModel.*.

CLASS ABL-XLSX-Writer.ABL-XLSX-Writer.writer.Styles INHERITS xmlwriter:
        
    DEFINE VARIABLE JSO_Style_Properties        AS CLASS JsonObject     NO-UNDO.

    DEFINE VARIABLE heading_pairs       AS CLASS JsonArray      NO-UNDO.
    DEFINE VARIABLE part_name           AS CLASS JsonArray      NO-UNDO.
    DEFINE VARIABLE JSO_patterns        AS CLASS JsonArray      NO-UNDO.
    DEFINE VARIABLE JSO_Format_Codes    AS CLASS JsonObject     NO-UNDO.

    DEFINE PRIVATE PROPERTY number_format_count     AS INTEGER              GET. SET.
    DEFINE PRIVATE PROPERTY font_count              AS INTEGER              GET. SET.
    DEFINE PRIVATE PROPERTY fill_count              AS INTEGER INITIAL 2    GET. SET.     
    DEFINE PRIVATE PROPERTY border_count            AS INTEGER INITIAL 2    GET. SET.     
    DEFINE PRIVATE PROPERTY xf_formats              AS HANDLE               GET. SET.   /** Handle to a temp-table of format Objects.**/  
    DEFINE PRIVATE PROPERTY fonts                   AS HANDLE               GET. SET.   /** Handle to fonts temp-table.**/  
    DEFINE PRIVATE PROPERTY Num_Format              AS HANDLE               GET. SET. 
    DEFINE PRIVATE PROPERTY Fills                   AS HANDLE               GET. SET.   /** Handle to Fills temp-table.**/  
    
    DEFINE TEMP-TABLE ttcellXfs NO-UNDO XML-NODE-NAME 'cellXfs'
        FIELD number_format_index AS INTEGER    XML-NODE-TYPE 'ATTRIBUTE' XML-NODE-NAME 'numFmtId' 
        FIELD font_index          AS INTEGER    XML-NODE-TYPE 'ATTRIBUTE' XML-NODE-NAME 'fontId' 
        FIELD fill_index          AS INTEGER    XML-NODE-TYPE 'ATTRIBUTE' XML-NODE-NAME 'fillId' 
        FIELD xfId                AS INTEGER    XML-NODE-TYPE 'ATTRIBUTE' XML-NODE-NAME 'xfId' 
        INDEX idxnumFmtId   IS PRIMARY UNIQUE 
            number_format_index.

    DEFINE TEMP-TABLE ttxf_formats NO-UNDO
        FIELD number_format_index AS INTEGER XML-NODE-TYPE 'ATTRIBUTE' XML-NODE-NAME 'numFmtId' 
        FIELD number_format       AS CHARACTER.

    DEFINE PRIVATE TEMP-TABLE ttxf_format NO-UNDO
            FIELD has_font          AS LOGICAL   
            FIELD has_fill          AS LOGICAL
            FIELD has_border        AS LOGICAL
            FIELD font_condense     AS LOGICAL
            FIELD font_extend       AS LOGICAL
            FIELD bold              AS LOGICAL
            FIELD font_strikeout    AS LOGICAL
            FIELD font_outline      AS LOGICAL
            FIELD font_shadow       AS LOGICAL
            FIELD alignment    AS INTEGER
            FIELD protection       AS INTEGER  /* superscript OR subscript*/
            .

    DEFINE VARIABLE data       AS CHARACTER   NO-UNDO.
    
    CONSTRUCTOR Styles():
        JSO_Style_Properties  = NEW JsonObject().
        heading_pairs = NEW JsonArray().
        part_name     = NEW JsonArray().

        JSO_Format_Codes = NEW JSONOBJECT().

        JSO_Format_Codes:ADD('0','General').
        JSO_Format_Codes:ADD('1','0').
        JSO_Format_Codes:ADD('2','0.00').
        JSO_Format_Codes:ADD('3','#,##0').
        JSO_Format_Codes:ADD('4','#,##0.00').
        JSO_Format_Codes:ADD('5','($#,##0_);($#,##0)').
        JSO_Format_Codes:ADD('6','($#,##0_);[Red]($#,##0)').
        JSO_Format_Codes:ADD('7','($#,##0.00_);($#,##0.00)').
        JSO_Format_Codes:ADD('8','($#,##0.00_);[Red]($#,##0.00)').
        JSO_Format_Codes:ADD('9','0%').
        JSO_Format_Codes:ADD('10','0.00%').
        JSO_Format_Codes:ADD('11','0.00E+00').
        JSO_Format_Codes:ADD('12','# ?/?').
        JSO_Format_Codes:ADD('13','# ??/??').
        JSO_Format_Codes:ADD('14','m/d/yy').
        JSO_Format_Codes:ADD('15','d-mmm-yy').
        JSO_Format_Codes:ADD('16','d-mmm').
        JSO_Format_Codes:ADD('17','mmm-yy').
        JSO_Format_Codes:ADD('18','h:mm AM/PM').
        JSO_Format_Codes:ADD('19','h:mm:ss AM/PM').
        JSO_Format_Codes:ADD('20','h:mm').
        JSO_Format_Codes:ADD('21','h:mm:ss').
        JSO_Format_Codes:ADD('22','m/d/yy h:mm').
        JSO_Format_Codes:ADD('23','(#,##0_);(#,##0)').
        JSO_Format_Codes:ADD('24','(#,##0_);[Red](#,##0)').
        JSO_Format_Codes:ADD('25','(#,##0.00_);(#,##0.00)').
        JSO_Format_Codes:ADD('26','(#,##0.00_);[Red](#,##0.00)').
        JSO_Format_Codes:ADD('27','_(* #,##0_);_(* (#,##0);_(* "-"_);_(_)').
        JSO_Format_Codes:ADD('28','_($* #,##0_);_($* (#,##0);_($* "-"_);_(_)').
        JSO_Format_Codes:ADD('29','_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(_)').
        JSO_Format_Codes:ADD('30','_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(_)').
        JSO_Format_Codes:ADD('31','mm:ss').
        JSO_Format_Codes:ADD('32','[h]:mm:ss').
        JSO_Format_Codes:ADD('33','mm:ss.0').
        JSO_Format_Codes:ADD('34','##0.0E+0').
        JSO_Format_Codes:ADD('35','@').

        JSO_patterns = NEW JSONARRAY().

        JSO_patterns:ADD('none').
        JSO_patterns:ADD('solid').
        JSO_patterns:ADD('mediumGray').
        JSO_patterns:ADD('darkGray').
        JSO_patterns:ADD('lightGray').
        JSO_patterns:ADD('darkHorizontal').
        JSO_patterns:ADD('darkVertical').
        JSO_patterns:ADD('darkDown').
        JSO_patterns:ADD('darkUp').
        JSO_patterns:ADD('darkGrid').
        JSO_patterns:ADD('darkTrellis').
        JSO_patterns:ADD('lightHorizontal').
        JSO_patterns:ADD('lightVertical').
        JSO_patterns:ADD('lightDown').
        JSO_patterns:ADD('lightUp').
        JSO_patterns:ADD('lightGrid').
        JSO_patterns:ADD('lightTrellis').
        JSO_patterns:ADD('gray125').
        JSO_patterns:ADD('gray0625').

    END CONSTRUCTOR.

    DESTRUCTOR Styles():
        DELETE OBJECT heading_pairs.
        DELETE OBJECT JSO_Style_Properties.
        DELETE OBJECT JSO_Format_Codes.
    END DESTRUCTOR.

    METHOD PUBLIC VOID _set_xf_formats(INPUT hnxf_formats AS HANDLE,
                                       INPUT hnfonts      AS HANDLE,
                                       INPUT hnNumFormat  AS HANDLE,
                                       INPUT hnFills      AS HANDLE):

        /** hnxf_formats is the handle to a temp-table. The Temp-table 
            contains a list of all the Format Class Objects.**/

        DEFINE VARIABLE hnQuery  AS HANDLE      NO-UNDO.
        DEFINE VARIABLE hnBuffer AS HANDLE      NO-UNDO.
        DEFINE VARIABLE obj_FormatInstance      AS CLASS format.

        IF VALID-HANDLE(hnxf_formats) AND hnxf_formats:TYPE EQ 'TEMP-TABLE' THEN
            THIS-OBJECT:xf_formats = hnxf_formats. 

        IF VALID-HANDLE(hnfonts) AND hnfonts:TYPE EQ 'TEMP-TABLE' THEN
            THIS-OBJECT:fonts = hnfonts. 

        IF VALID-HANDLE(hnNumFormat) AND hnNumFormat:TYPE EQ 'TEMP-TABLE' THEN
            THIS-OBJECT:Num_Format = hnNumFormat. 
        
        IF VALID-HANDLE(hnFills) AND hnFills:TYPE EQ 'TEMP-TABLE' THEN
            THIS-OBJECT:Fills = hnFills. 



    END METHOD.

    METHOD PUBLIC VOID _assemble_xml_file():

        THIS-OBJECT:_write_style_sheet().

        THIS-OBJECT:_xml_close().
        
    END METHOD.

    METHOD PRIVATE VOID _write_style_sheet():
        
        hnSAXWriter:START-ELEMENT('styleSheet').
        hnSAXWriter:DECLARE-NAMESPACE('http://schemas.openxmlformats.org/spreadsheetml/2006/main').

        THIS-OBJECT:_write_num_fmts().     

        THIS-OBJECT:_write_fonts().

        THIS-OBJECT:_write_fills().

        THIS-OBJECT:_write_borders().

        THIS-OBJECT:_write_cell_style_xfs().

        THIS-OBJECT:_write_cell_xfs().

        THIS-OBJECT:_write_cell_styles().

        THIS-OBJECT:_write_table_styles().

        THIS-OBJECT:_write_colors().
  
        hnSAXWriter:END-ELEMENT('styleSheet').
    END METHOD.

    METHOD VOID _write_num_fmts():
        

        DEFINE VARIABLE hnNumFmtsBuffer AS HANDLE      NO-UNDO.
        DEFINE VARIABLE hnQuery         AS HANDLE      NO-UNDO.
        DEFINE VARIABLE hnNum_Format    AS HANDLE      NO-UNDO.
        
        hnNum_Format = THIS-OBJECT:Num_Format.

        IF NOT VALID-HANDLE(hnNum_Format) OR
           NOT hnNum_Format:TYPE EQ 'TEMP-TABLE' THEN
            RETURN.

        /* If not records exist, leave. */
        IF NOT hnNum_Format:HAS-RECORDS THEN
            RETURN.

        hnNumFmtsBuffer = hnNum_Format:DEFAULT-BUFFER-HANDLE.

        CREATE QUERY hnQuery.

        hnQuery:ADD-BUFFER(hnNumFmtsBuffer). 
        hnQuery:QUERY-PREPARE("PRESELECT EACH ttNumFormats BY ttNumFormats.number_format_index").
        hnQuery:QUERY-OPEN().

        IF hnQuery:IS-OPEN THEN
        DO:
            hnSAXWriter:START-ELEMENT('numFmts').
            hnSAXWriter:INSERT-ATTRIBUTE('count' , STRING(  hnQuery:NUM-RESULTS )). /** NOTE: NUM-RESULTS only works if you use 'PRESELECT' rather 'EACH'. **/
            
            hnQuery:GET-FIRST() NO-ERROR.                       
            REPEAT WHILE NOT hnQuery:QUERY-OFF-END:
    
                THIS-OBJECT:_write_num_fmt(INPUT hnNumFmtsBuffer::number_format_index,
                                           INPUT hnNumFmtsBuffer::format_code ).
    
                hnQuery:GET-NEXT() NO-ERROR.                        
            END.
    
            hnQuery:QUERY-CLOSE().

            hnSAXWriter:END-ELEMENT('numFmts').
        END.
        DELETE OBJECT hnQuery.

    END METHOD.

    METHOD VOID _write_num_fmt(INPUT num_Format_Id   AS INTEGER,
                               INPUT num_Format_Code AS CHARACTER):

        DEFINE VARIABLE ch_FormatCode AS CHARACTER   NO-UNDO INITIAL 'General'.
        
        IF num_Format_Id LT 164 THEN
        DO:
           IF THIS-OBJECT:JSO_Format_Codes:HAS( STRING(num_Format_Id) ) THEN
               ch_FormatCode = THIS-OBJECT:JSO_Format_Codes:GETCHARACTER(STRING(num_Format_Id)).
           ELSE
               ch_FormatCode = 'General'.
        END.
        ELSE
            ch_FormatCode = num_Format_Code.
           
        hnSAXWriter:WRITE-EMPTY-ELEMENT('numFmt').
        hnSAXWriter:INSERT-ATTRIBUTE('numFmtId' , STRING( num_Format_Id )).
        hnSAXWriter:INSERT-ATTRIBUTE('formatCode' , ch_FormatCode).
            

    END METHOD.
    
    METHOD VOID _write_fonts():

        DEFINE VARIABLE hnBuffer AS HANDLE      NO-UNDO.
        DEFINE VARIABLE hnQuery  AS HANDLE      NO-UNDO.

        IF NOT THIS-OBJECT:fonts:HAS-RECORDS THEN
        DO:
            hnSAXWriter:WRITE-EMPTY-ELEMENT('fonts').
            hnSAXWriter:INSERT-ATTRIBUTE('count','0').
            RETURN.
        END.

        hnBuffer = THIS-OBJECT:fonts:DEFAULT-BUFFER-HANDLE.

        CREATE QUERY hnQuery.

        hnQuery:ADD-BUFFER(hnBuffer).
        hnQuery:QUERY-PREPARE('PRESELECT EACH ttFonts BY ttFonts.font_index').
        hnQuery:QUERY-OPEN().

        IF hnQuery:QUERY-OPEN THEN
        DO:
            hnSAXWriter:START-ELEMENT('fonts').
            hnSAXWriter:INSERT-ATTRIBUTE('count', STRING(hnQuery:NUM-RESULTS) ).
            
            hnQuery:GET-FIRST() NO-ERROR.
    
            REPEAT WHILE NOT hnQuery:QUERY-OFF-END:
    
                THIS-OBJECT:_write_font(INPUT hnBuffer).
                hnQuery:GET-NEXT() NO-ERROR.
            END.
    
            hnSAXWriter:END-ELEMENT('fonts').
            hnQuery:QUERY-CLOSE().
        END.
        DELETE OBJECT hnQuery.

    END METHOD.

    METHOD VOID _write_font (INPUT hnFontBuffer AS HANDLE):
                    
/*         /** Write  the Default font.. **/               */
/*         hnSAXWriter:START-ELEMENT('font').              */
/*         hnSAXWriter:WRITE-EMPTY-ELEMENT('sz').          */
/*         hnSAXWriter:INSERT-ATTRIBUTE('val', '11').      */
/*                                                         */
/*         hnSAXWriter:WRITE-EMPTY-ELEMENT('color').       */
/*         hnSAXWriter:INSERT-ATTRIBUTE('theme', '1').     */
/*                                                         */
/*         hnSAXWriter:WRITE-EMPTY-ELEMENT('name').        */
/*         hnSAXWriter:INSERT-ATTRIBUTE('val', 'Calibri'). */
/*                                                         */
/*         hnSAXWriter:WRITE-EMPTY-ELEMENT('family').      */
/*         hnSAXWriter:INSERT-ATTRIBUTE('val', '2').       */
/*                                                         */
/*         hnSAXWriter:WRITE-EMPTY-ELEMENT('scheme').      */
/*         hnSAXWriter:INSERT-ATTRIBUTE('val', 'minor').   */
/*         hnSAXWriter:END-ELEMENT('font').                */
        
        IF NOT VALID-HANDLE( hnFontBuffer) OR 
           NOT hnFontBuffer:type EQ 'BUFFER' THEN
            RETURN.

        hnSAXWriter:START-ELEMENT('font').

        /* bold */
        IF hnFontBuffer::bold THEN
        DO:
            hnSAXWriter:WRITE-EMPTY-ELEMENT('b').
            hnSAXWriter:INSERT-ATTRIBUTE('val', '1').
        END.

        /* italic */
        IF hnFontBuffer::italic THEN
        DO:
            hnSAXWriter:WRITE-EMPTY-ELEMENT('i').
            hnSAXWriter:INSERT-ATTRIBUTE('val', '1').
        END.

        /* underline */
        IF LOGICAL(hnFontBuffer::f_underline) THEN
            THIS-OBJECT:_write_underline( hnFontBuffer::f_underline ).
        
/*         hnSAXWriter:INSERT-ATTRIBUTE('x14ac:knownFonts','1'). */
                
        hnSAXWriter:WRITE-EMPTY-ELEMENT('sz').
        hnSAXWriter:INSERT-ATTRIBUTE('val', STRING(hnFontBuffer::font_size)).
        hnSAXWriter:WRITE-EMPTY-ELEMENT('color').

        IF hnFontBuffer::colour EQ '' THEN
            hnSAXWriter:INSERT-ATTRIBUTE('theme', '1').
        ELSE
        DO:
            DEFINE VARIABLE chTempColour AS CHARACTER   NO-UNDO.
            IF hnFontBuffer::colour BEGINS '#' THEN
            DO:
                chTempColour = REPLACE(hnFontBuffer::colour,'#','FF').
                chTempColour = CAPS(chTempColour).
                hnSAXWriter:INSERT-ATTRIBUTE('rgb', chTempColour).
            END.
        END.

        hnSAXWriter:WRITE-EMPTY-ELEMENT('name').
        hnSAXWriter:INSERT-ATTRIBUTE('val', STRING(hnFontBuffer::font_name)).
        hnSAXWriter:WRITE-EMPTY-ELEMENT('family').
        hnSAXWriter:INSERT-ATTRIBUTE('val', STRING(hnFontBuffer::family)).
        hnSAXWriter:WRITE-EMPTY-ELEMENT('scheme').
        hnSAXWriter:INSERT-ATTRIBUTE('val', 'minor').
        hnSAXWriter:END-ELEMENT('font').

    END METHOD.

    METHOD PRIVATE VOID _write_underline(INPUT IN_underline AS INTEGER):
        
        hnSAXWriter:WRITE-EMPTY-ELEMENT('u').

        /**  doubleAccounting,singleAccounting,double **/

        CASE in_underline:
            WHEN 1 THEN
                .     /** Do nothing. No Attribute Required. **/
            WHEN 2 THEN
                hnSAXWriter:INSERT-ATTRIBUTE('val', 'double'). /** Double Line.. **/
            WHEN 3 THEN
                hnSAXWriter:INSERT-ATTRIBUTE('val', 'singleAccounting'). /** Single Line Accounting.. **/
            WHEN 4 THEN
                hnSAXWriter:INSERT-ATTRIBUTE('val', 'doubleAccounting'). /** Double Line Accounting.. **/
            
        END CASE.

    END METHOD.

    
    METHOD VOID _write_fills():

        DEFINE VARIABLE hnFills       AS HANDLE      NO-UNDO.
        DEFINE VARIABLE hnFill       AS HANDLE      NO-UNDO.
        DEFINE VARIABLE hnQuery          AS HANDLE      NO-UNDO.
        DEFINE VARIABLE obFormatInstance AS CLASS FORMAT.

        hnFills = THIS-OBJECT:Fills.

        IF NOT VALID-HANDLE(hnFills ) OR
           NOT hnFills :TYPE EQ 'TEMP-TABLE' THEN
            RETURN.

        IF NOT hnFills:HAS-RECORDS THEN
        DO:
            hnSAXWriter:START-ELEMENT('fills').
            hnSAXWriter:INSERT-ATTRIBUTE('count', '2' ).

            /** Write the default fill element **/
            THIS-OBJECT:_write_default_fill(INPUT 'none').
            THIS-OBJECT:_write_default_fill(INPUT 'gray125').

            hnSAXWriter:END-ELEMENT('fills').
            RETURN.
        END.

        hnFill = hnFills:DEFAULT-BUFFER-HANDLE.

        CREATE QUERY hnQuery.
        hnQuery:ADD-BUFFER(hnFill).
        hnQuery:QUERY-PREPARE('PRESELECT EACH ttFills BY ttFills.fill_index').
        hnQuery:QUERY-OPEN().

        hnSAXWriter:START-ELEMENT('fills').
        hnSAXWriter:INSERT-ATTRIBUTE('count', STRING( hnQuery:NUM-RESULTS ) ).
        
        IF hnQuery:IS-OPEN THEN
        DO:
            
            hnQuery:GET-FIRST() NO-ERROR.

            REPEAT WHILE NOT hnQuery:QUERY-OFF-END:
                THIS-OBJECT:_write_fill(INPUT hnFill). /** Pass the buffer handle**/
                hnQuery:GET-NEXT() NO-ERROR.
            END.

            hnQuery:QUERY-CLOSE().
        END.
        
        DELETE OBJECT hnQuery.

        hnSAXWriter:END-ELEMENT('fills').

    END METHOD.

    METHOD PRIVATE VOID _write_default_fill(INPUT pattern_type AS CHARACTER):

        hnSAXWriter:START-ELEMENT('fill').
        hnSAXWriter:WRITE-EMPTY-ELEMENT('patternFill').
        hnSAXWriter:INSERT-ATTRIBUTE('patternType', pattern_type ).
        hnSAXWriter:END-ELEMENT('fill').

    END METHOD.

    METHOD PRIVATE VOID _write_fill(INPUT hnFill AS HANDLE):

        DEFINE VARIABLE inPattern   AS INTEGER      NO-UNDO.
        DEFINE VARIABLE chbg_colour AS CHARACTER    NO-UNDO.
        DEFINE VARIABLE chfg_colour AS CHARACTER    NO-UNDO.

        ASSIGN
            inPattern    = hnFill::pattern
            chbg_colour  = hnFill::bg_colour   
            chfg_colour  = hnFill::fg_colour.

        hnSAXWriter:START-ELEMENT('fill').
        hnSAXWriter:START-ELEMENT('patternFill').
        
        IF inPattern GT 0 THEN
            hnSAXWriter:INSERT-ATTRIBUTE('patternType', JSO_patterns:GetCharacter( inPattern ) ).

        IF chfg_colour NE '' THEN
        DO:
            chfg_colour = REPLACE(chfg_colour,'#','FF').
            hnSAXWriter:WRITE-EMPTY-ELEMENT('fgColor').
            hnSAXWriter:INSERT-ATTRIBUTE('rgb', CAPS(chfg_colour) ).
        END.

        IF chbg_colour NE '' THEN
        DO:
            chbg_colour = REPLACE(chbg_colour,'#','FF').
            hnSAXWriter:WRITE-EMPTY-ELEMENT('bgColor').
            hnSAXWriter:INSERT-ATTRIBUTE('rgb', CAPS(chbg_colour) ).
        END.

        hnSAXWriter:END-ELEMENT('patternFill').
        hnSAXWriter:END-ELEMENT('fill').
    END METHOD.
    
    METHOD VOID _write_borders():


        hnSAXWriter:START-ELEMENT('borders').
        hnSAXWriter:INSERT-ATTRIBUTE('count', '1').

        THIS-OBJECT:_write_border().

        hnSAXWriter:END-ELEMENT('borders').
    
/*         hnSAXWriter:START-ELEMENT('borders').                                     */
/*         hnSAXWriter:INSERT-ATTRIBUTE('count', STRING(THIS-OBJECT:border_count) ). */
/*                                                                                   */
/*                                                                                   */
/*         FOR EACH ttxf_format                                                      */
/*             WHERE ttxf_format.has_border:                                         */
/*             THIS-OBJECT:_write_border().                                          */
/*         END.                                                                      */
/*                                                                                   */
/*         hnSAXWriter:END-ELEMENT('borders').                                       */

    END METHOD.

    METHOD VOID _write_border():
        
        hnSAXWriter:START-ELEMENT('border').
    
            hnSAXWriter:WRITE-EMPTY-ELEMENT('left').
            hnSAXWriter:WRITE-EMPTY-ELEMENT('right').
            hnSAXWriter:WRITE-EMPTY-ELEMENT('top').
            hnSAXWriter:WRITE-EMPTY-ELEMENT('bottom').
            hnSAXWriter:WRITE-EMPTY-ELEMENT('diagonal').
    
        hnSAXWriter:END-ELEMENT('border').

    END METHOD.

    METHOD VOID _write_cell_style_xfs():
        /* # Write the <cellStyleXfs> element. */
        hnSAXWriter:START-ELEMENT('cellStyleXfs').
        hnSAXWriter:INSERT-ATTRIBUTE('count', '1').

        THIS-OBJECT:_write_style_xf().

        hnSAXWriter:END-ELEMENT('cellStyleXfs').
    END METHOD.

    METHOD PRIVATE VOID _write_style_xf():
        hnSAXWriter:WRITE-EMPTY-ELEMENT('xf').
        hnSAXWriter:INSERT-ATTRIBUTE('numFmtId', '0').
        hnSAXWriter:INSERT-ATTRIBUTE('fontId', '0').
        hnSAXWriter:INSERT-ATTRIBUTE('fillId', '0').
        hnSAXWriter:INSERT-ATTRIBUTE('borderId', '0').
    END METHOD.

    METHOD PRIVATE VOID _write_cell_xfs():
        /* # Write the style <xf> element. */

        DEFINE VARIABLE hnxf_formats        AS HANDLE      NO-UNDO.
        DEFINE VARIABLE hnxf_format         AS HANDLE      NO-UNDO.
        DEFINE VARIABLE hnQuery             AS HANDLE      NO-UNDO.
        DEFINE VARIABLE obFormatInstance    AS CLASS FORMAT NO-UNDO.
        
        hnxf_formats = THIS-OBJECT:xf_formats.

        IF NOT VALID-HANDLE(xf_formats) OR
           NOT xf_formats:TYPE EQ 'TEMP-TABLE' THEN
            RETURN.

        hnxf_format = hnxf_formats:DEFAULT-BUFFER-HANDLE.

        CREATE QUERY hnQuery.
        hnQuery:ADD-BUFFER(hnxf_format).
        hnQuery:QUERY-PREPARE('PRESELECT EACH ttxf_Formats BREAK BY ttxf_Formats.format_index').
        hnQuery:QUERY-OPEN().

        IF hnQuery:IS-OPEN THEN
        DO:
            
            hnSAXWriter:START-ELEMENT('cellXfs').
            hnSAXWriter:INSERT-ATTRIBUTE('count', STRING( hnQuery:NUM-RESULTS ) ).
            
            hnQuery:GET-FIRST() NO-ERROR.

/*             hnSAXWriter:WRITE-COMMENT('ttxf_Formats.format_index: ' + STRING( hnxf_format::format_index ) ). */

            REPEAT WHILE NOT hnQuery:QUERY-OFF-END:

                IF hnQuery:FIRST-OF ( 1 ) THEN
                DO:
                    obFormatInstance = hnxf_format::ObjFormat.
                    THIS-OBJECT:_write_xf(INPUT obFormatInstance).
                END.
                hnQuery:GET-NEXT() NO-ERROR.
            END.

            hnQuery:QUERY-CLOSE().

            hnSAXWriter:END-ELEMENT('cellXfs').

        END.
        
        DELETE OBJECT hnQuery.

    END METHOD.

    METHOD PRIVATE VOID _write_xf(INPUT ob_Format_Instance AS CLASS FORMAT):
        
        DEFINE VARIABLE hn_alignment_attributes AS HANDLE.

        IF NOT VALID-OBJECT(ob_Format_Instance) THEN
            RETURN.

        /* Write XF with sub-elements if required. */
        hn_alignment_attributes = ob_Format_Instance:_get_align_properties().

        IF hn_alignment_attributes:NUM-ITEMS GT 0 THEN
            hnSAXWriter:START-ELEMENT('xf').
        ELSE
            hnSAXWriter:WRITE-EMPTY-ELEMENT('xf').
    
/*         hnSAXWriter:INSERT-ATTRIBUTE('Id',  STRING(ob_Format_Instance:xf_index) ). */

        hnSAXWriter:INSERT-ATTRIBUTE('numFmtId', STRING(ob_Format_Instance:num_format_index) ).
        hnSAXWriter:INSERT-ATTRIBUTE('fontId',   STRING(ob_Format_Instance:font_index) ).
        hnSAXWriter:INSERT-ATTRIBUTE('fillId',   STRING(ob_Format_Instance:fill_index) ).
        hnSAXWriter:INSERT-ATTRIBUTE('borderId', STRING(ob_Format_Instance:border_index) ).

        IF LOGICAL(ob_Format_Instance:num_format_index) THEN
            hnSAXWriter:INSERT-ATTRIBUTE('applyNumberFormat', '1' ).

        IF LOGICAL(ob_Format_Instance:font_index) THEN
            hnSAXWriter:INSERT-ATTRIBUTE('applyFont', '1' ).

        IF LOGICAL(ob_Format_Instance:fill_index) THEN
            hnSAXWriter:INSERT-ATTRIBUTE('applyFill', '1' ).

        IF LOGICAL(ob_Format_Instance:border_index) THEN
            hnSAXWriter:INSERT-ATTRIBUTE('applyBorder', '1' ).

        IF hn_alignment_attributes:NUM-ITEMS GT 0 THEN
        DO:
            /** Check if XF format has alignment properties set. **/
            hnSAXWriter:INSERT-ATTRIBUTE('applyAlignment', '1' ).
            hnSAXWriter:WRITE-EMPTY-ELEMENT('alignment',?, hn_alignment_attributes).
            DELETE OBJECT hn_alignment_attributes.
            hnSAXWriter:END-ELEMENT('xf').
        END.
        
    END METHOD.



    METHOD PRIVATE VOID _write_cell_styles():
         /*# Write the <cellStyles> element.*/
        hnSAXWriter:START-ELEMENT('cellStyles').
        hnSAXWriter:INSERT-ATTRIBUTE('count', '1').
        THIS-OBJECT:_write_cell_style().
        hnSAXWriter:END-ELEMENT('cellStyles').
    END METHOD.

    METHOD PRIVATE VOID _write_cell_style():
         /*# Write the <cellStyle> element.*/
        hnSAXWriter:WRITE-EMPTY-ELEMENT('cellStyle').  
        hnSAXWriter:INSERT-ATTRIBUTE('name', 'Normal').
        hnSAXWriter:INSERT-ATTRIBUTE('xfId', '0').
        hnSAXWriter:INSERT-ATTRIBUTE('builtinId', '0').
    END METHOD.

    METHOD VOID _write_table_styles():
        hnSAXWriter:WRITE-EMPTY-ELEMENT('tableStyles').
        hnSAXWriter:INSERT-ATTRIBUTE('count', '0' ).
        hnSAXWriter:INSERT-ATTRIBUTE('defaultTableStyle', 'TableStyleMedium2' ).
        hnSAXWriter:INSERT-ATTRIBUTE('defaultPivotStyle', 'PivotStyleLight16' ).
    END METHOD.

    METHOD VOID _write_colors():
    END METHOD.

END CLASS.



