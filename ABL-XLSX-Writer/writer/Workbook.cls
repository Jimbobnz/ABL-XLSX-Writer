
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

CLASS ABL-XLSX-Writer.ABL-XLSX-Writer.writer.Workbook INHERITS xmlwriter:

    DEFINE PUBLIC PROPERTY filename            AS CHARACTER                    GET. PRIVATE SET.
    DEFINE PRIVATE PROPERTY TempDir             AS CHARACTER                    GET. SET.
    DEFINE PRIVATE PROPERTY sheetname           AS CHARACTER    INITIAL 'Sheet' GET. SET.
    DEFINE PRIVATE PROPERTY sheetname_count     AS INTEGER      INITIAL 0       GET. SET.
    DEFINE PRIVATE PROPERTY num_format_count    AS INTEGER      INITIAL 0       GET. SET.
    DEFINE PRIVATE PROPERTY format_count        AS INTEGER      INITIAL 0       GET. SET.
    DEFINE PUBLIC  PROPERTY SharedStringTable   AS CLASS SharedStringTable      GET. SET.
    DEFINE PUBLIC  PROPERTY xf_formats          AS HANDLE                       GET. PRIVATE SET.
    
    

    DEFINE PRIVATE PROPERTY xf_Format_indices   AS CLASS JSONOBJECT             GET. SET.
    DEFINE PRIVATE PROPERTY dxf_Format_indices  AS CLASS JSONOBJECT             GET. SET.

    DEFINE PRIVATE PROPERTY obj_DefaultFormats  AS CLASS format                 GET. SET.
    
    

    DEFINE PRIVATE TEMP-TABLE ttWorkSheet NO-UNDO
        FIELD SheetName     AS CHARACTER
        FIELD SheetID       AS INTEGER
        FIELD HIDDEN        AS LOGICAL INITIAL FALSE
        FIELD ObjWorkSheet  AS CLASS Progress.Lang.OBJECT 
        INDEX idxSheetID IS PRIMARY
            SheetID .

    DEFINE PRIVATE TEMP-TABLE ttxf_Formats NO-UNDO
        FIELD format_index  AS INTEGER
        FIELD ObjFormat     AS CLASS Progress.Lang.OBJECT
        INDEX idxIndex IS PRIMARY 
            format_index
        INDEX idxFormat 
            ObjFormat .

    DEFINE PRIVATE TEMP-TABLE ttProperties NO-UNDO 
        FIELD wb_title   AS CHARACTER    SERIALIZE-NAME 'title'
        FIELD subject    AS CHARACTER    SERIALIZE-NAME 'subject'
        FIELD author     AS CHARACTER    SERIALIZE-NAME 'author'
        FIELD manager    AS CHARACTER    SERIALIZE-NAME 'manager'
        FIELD company    AS CHARACTER    SERIALIZE-NAME 'company'
        FIELD keywords   AS CHARACTER    SERIALIZE-NAME 'keywords'
        FIELD comments   AS CHARACTER    SERIALIZE-NAME 'comments'
        FIELD category   AS CHARACTER    SERIALIZE-NAME 'category'
        FIELD wb_status  AS CHARACTER    SERIALIZE-NAME 'status'
        .

    DEFINE TEMP-TABLE ttNumFormats NO-UNDO XML-NODE-NAME 'numFmts'
        FIELD number_format_index AS INTEGER    XML-NODE-TYPE 'ATTRIBUTE' XML-NODE-NAME 'numFmtId' 
        FIELD format_code         AS CHARACTER  XML-NODE-TYPE 'ATTRIBUTE' XML-NODE-NAME 'formatCode'
        INDEX idxnumFmtId   IS PRIMARY  
            number_format_index
        INDEX idxFormatCode IS UNIQUE 
            format_code.

    DEFINE TEMP-TABLE ttFonts NO-UNDO XML-NODE-NAME 'font'
        FIELD font_index    AS INTEGER    XML-NODE-TYPE 'HIDDEN' 
        FIELD font_name     AS CHARACTER  XML-NODE-TYPE 'ELEMENT' XML-NODE-NAME 'name' 
        FIELD font_size     AS DECIMAL    XML-NODE-TYPE 'ELEMENT' XML-NODE-NAME 'sz' 
        FIELD charset       AS INTEGER    XML-NODE-TYPE 'ELEMENT' XML-NODE-NAME 'charset' 
        FIELD family        AS INTEGER    XML-NODE-TYPE 'ELEMENT' XML-NODE-NAME 'family' 
        FIELD bold          AS LOGICAL    XML-NODE-TYPE 'ELEMENT' XML-NODE-NAME 'b'
        FIELD italic        AS LOGICAL    XML-NODE-TYPE 'ELEMENT' XML-NODE-NAME 'i'
        FIELD f_underline   AS INTEGER    XML-NODE-TYPE 'ELEMENT' XML-NODE-NAME 'u'
        FIELD strike        AS LOGICAL    XML-NODE-TYPE 'ELEMENT' XML-NODE-NAME 'strike'
        FIELD outline       AS LOGICAL    XML-NODE-TYPE 'ELEMENT' XML-NODE-NAME 'outline'
        FIELD shadow        AS LOGICAL    XML-NODE-TYPE 'ELEMENT' XML-NODE-NAME 'shadow'
        FIELD condense      AS LOGICAL    XML-NODE-TYPE 'ELEMENT' XML-NODE-NAME 'condense'
        FIELD extend        AS LOGICAL    XML-NODE-TYPE 'ELEMENT' XML-NODE-NAME 'extend'
        FIELD colour        AS CHARACTER  XML-NODE-TYPE 'ELEMENT' XML-NODE-NAME 'color'     /** Store a JsonObject **/
        INDEX idxFontIndex IS PRIMARY UNIQUE
            font_index.

    
    CONSTRUCTOR Workbook(INPUT pchExcel AS CHARACTER):
        
        /** Check for logical file locks created be MS Office Excel**/

        IF SEARCH('~~$' + pchExcel) NE ?  THEN
        DO:
            MESSAGE SUBSTITUTE('Microsoft Excel already has this "&1" open.',
                               pchExcel)
                VIEW-AS ALERT-BOX ERROR.

            RETURN ERROR.
        END.

        THIS-OBJECT:filename           = pchExcel.
        THIS-OBJECT:TempDir            = SESSION:TEMP-DIR.
        THIS-OBJECT:SharedStringTable  = NEW SharedStringTable().

        THIS-OBJECT:xf_Format_indices  = NEW JSONOBJECT(). 
        THIS-OBJECT:dxf_Format_indices = NEW JSONOBJECT().

        THIS-OBJECT:xf_formats         = TEMP-TABLE ttxf_Formats:HANDLE.
        
        /** Load the defaults Formats etc. **/
        THIS-OBJECT:add_Format().
        
    END CONSTRUCTOR.

    METHOD PUBLIC HANDLE xf_formats():
        RETURN TEMP-TABLE ttxf_Formats:HANDLE.
    END METHOD.

    METHOD PUBLIC HANDLE Fonts():
        RETURN TEMP-TABLE ttFonts:HANDLE.
    END METHOD.

    METHOD PUBLIC HANDLE Num_Format():
        RETURN TEMP-TABLE ttNumFormats:HANDLE.
    END METHOD.

    
    


    /** PUBLIC METHODS... **/

    METHOD PUBLIC VOID set_properties(INPUT ch_ParameterKey   AS CHARACTER,
                                      INPUT ch_ParameterValue AS CHARACTER):
        
        DEFINE VARIABLE hnProperties    AS HANDLE      NO-UNDO.
        DEFINE VARIABLE hnPropertyField AS HANDLE      NO-UNDO.
        
        ch_ParameterKey = TRIM(ch_ParameterKey).            
        hnProperties    = TEMP-TABLE ttProperties:DEFAULT-BUFFER-HANDLE.

        hnProperties:FIND-FIRST() NO-ERROR.

        IF NOT hnProperties:AVAILABLE THEN
            hnProperties:BUFFER-CREATE.

        /** Overide to handle progress key words **/
        CASE ch_ParameterKey:
            WHEN 'title' THEN
                ch_ParameterKey = 'wb_title'.
            WHEN 'status' THEN
                ch_ParameterKey = 'wb_status'.
        END.

        hnPropertyField = hnProperties:BUFFER-FIELD(ch_ParameterKey).

        /** If a invalid ParameterKey is provided this will handle it. **/
        IF VALID-HANDLE(hnPropertyField) THEN
            hnPropertyField:BUFFER-VALUE() = ch_ParameterValue.

        hnProperties:BUFFER-RELEASE().

    END METHOD.

    METHOD PUBLIC CLASS worksheet add_worksheet():
        RETURN THIS-OBJECT:add_worksheet(INPUT THIS-OBJECT:sheetname + STRING(THIS-OBJECT:sheetname_count + 1) ).
    END METHOD.

    METHOD PUBLIC CLASS worksheet add_worksheet( INPUT pcSheetName AS CHARACTER):

        DEFINE VARIABLE ObjNewWorkSheet AS CLASS worksheet.
        
        THIS-OBJECT:sheetname_count = THIS-OBJECT:sheetname_count + 1.
        
        ObjNewWorkSheet = NEW Worksheet(INPUT pcSheetName).
        
        /** Share the Global Shared String Table Class Object. In Theory 
            there should only ever be one instance of the class object. **/
        ObjNewWorkSheet:SharedStringTable = THIS-OBJECT:SharedStringTable.

        CREATE ttWorkSheet.
        
        ASSIGN
            ttWorkSheet.SheetName    = pcSheetName
            ttWorkSheet.SheetID      = THIS-OBJECT:sheetname_count
            ttWorkSheet.ObjWorkSheet = ObjNewWorkSheet. 
        
        RETURN ObjNewWorkSheet.

    END METHOD.

    METHOD PUBLIC VOID close():
        THIS-OBJECT:_store_workbook().
    END METHOD.

    METHOD PUBLIC HANDLE worksheets():
            
        /** Return the temp-table's HANDLE attribute. **/
        IF VALID-HANDLE(TEMP-TABLE ttWorkSheet:HANDLE) THEN
           RETURN TEMP-TABLE ttWorkSheet:HANDLE.

    END METHOD.

    METHOD PUBLIC HANDLE properties():

        IF VALID-HANDLE(TEMP-TABLE ttProperties:HANDLE) THEN
           RETURN TEMP-TABLE ttProperties:HANDLE.

    END METHOD.

    METHOD PUBLIC CLASS Format add_Format():

        DEFINE VARIABLE ob_Format AS CLASS Format no-undo.

        
        ob_Format = NEW Format(INPUT THIS-OBJECT:xf_Format_indices,
                               INPUT THIS-OBJECT:dxf_Format_indices).

        THIS-OBJECT:format_count = THIS-OBJECT:format_count + 1.

        CREATE ttxf_Formats.

        ASSIGN
             ttxf_Formats.format_index = THIS-OBJECT:format_count
             ttxf_Formats.OBJFormat    = ob_Format.

        RETURN ob_Format.
    END.

    METHOD PRIVATE VOID _prepare_format_properties():

        MESSAGE '_prepare_formats()'.    
        THIS-OBJECT:_prepare_formats().
        MESSAGE '_prepare_fonts()'.    
        THIS-OBJECT:_prepare_fonts().
        MESSAGE '_prepare_num_formats()'.    
        THIS-OBJECT:_prepare_num_formats(). /** Number format/masks**/
        MESSAGE '_prepare_borders()'.    
        THIS-OBJECT:_prepare_borders().
        THIS-OBJECT:_prepare_fills().
    END METHOD.

    METHOD PRIVATE VOID _prepare_formats():

          /* Rearrange formats into index order.*/

        DEFINE VARIABLE in_index            AS INTEGER     NO-UNDO.
        DEFINE VARIABLE ob_format_instance  AS CLASS FORMAT NO-UNDO.
            

        DEFINE BUFFER ttxf_Formats FOR TEMP-TABLE ttxf_Formats.
        DEFINE BUFFER xf_Formats FOR TEMP-TABLE ttxf_Formats.

        FOR EACH ttxf_Formats:

            ob_format_instance = CAST(ttxf_Formats.OBJFormat, format).
            in_index = ob_format_instance:xf_index.

            ASSIGN
                ttxf_Formats.format_index = in_index.

        END.
    
    END METHOD.

    METHOD PRIVATE VOID _prepare_fonts():
        
        DEFINE VARIABLE ob_Format        AS CLASS format.
        DEFINE VARIABLE obAvaialbleFonts AS CLASS JsonObject.
        DEFINE VARIABLE inFont_Index     AS INTEGER      NO-UNDO.
        DEFINE VARIABLE chFontKey        AS CHARACTER    NO-UNDO.

        inFont_Index = 0.

        /** Create a font reocrd. **/
        obAvaialbleFonts      = NEW JsonObject().
        
        FOR EACH ttxf_Formats:
             ob_Format = CAST(ttxf_Formats.ObjFormat, format).

             chFontKey = ob_Format:_get_font_key().

             /** If the format object font as already been defined 
                 by another format class object, use that. **/

             IF obAvaialbleFonts:HAS(chFontKey) THEN
             DO:
                ob_Format:font_index = obAvaialbleFonts:GetInteger(chFontKey).
                ob_Format:has_font   = FALSE.
             END.
             ELSE
             DO:
                 /* This is a new font */
                 obAvaialbleFonts:ADD(chFontKey, inFont_Index).

                 ob_Format:font_index = inFont_Index.
                 ob_Format:has_font   = TRUE.
                 inFont_Index         = inFont_Index + 1. 

                 CREATE ttFonts.

                 ASSIGN
                     ttFonts.font_index  = ob_Format:font_index
                     ttFonts.font_name   = ob_Format:font_name
                     ttFonts.font_size   = ob_Format:font_size
                     ttFonts.charset     = ob_Format:font_charset
                     ttFonts.family      = ob_Format:font_family
                     ttFonts.bold        = ob_Format:bold
                     ttFonts.italic      = ob_Format:italic
                     ttFonts.f_underline = ob_Format:underline   /* doubleAccounting,singleAccounting,double*/     
                     ttFonts.strike      = ob_Format:font_strike
                     ttFonts.outline     = ob_Format:font_outline
                     ttFonts.shadow      = ob_Format:font_shadow 
                     ttFonts.colour      = ob_Format:font_colour.    /** NOTE TO SELF: color is a progress researce word.*/

             END.
        END.
        
    END METHOD.

    METHOD PRIVATE VOID _prepare_num_formats():

        DEFINE VARIABLE in_Index            AS INTEGER INITIAL 164  NO-UNDO.
        DEFINE VARIABLE ob_Format           AS CLASS format.
        DEFINE VARIABLE in_num_format_count AS INTEGER INITIAL 0    NO-UNDO.
        DEFINE VARIABLE ch_FormatCode       AS CHARACTER            NO-UNDO.
        DEFINE VARIABLE JSO_Num_Formats     AS CLASS JSONOBJECT     NO-UNDO.
        
        JSO_Num_Formats = NEW JSONOBJECT().
        
        FOR EACH ttxf_Formats:
            ob_Format = CAST(ttxf_Formats.ObjFormat, format).
        
            IF NOT VALID-OBJECT(ob_Format) THEN
                NEXT.
        
            ch_FormatCode = ob_Format:num_format.
            
/*             ob_Format:num_format_index = INTEGER(ch_FormatCode) NO-ERROR. */

            /** If the format object font as already been defined 
                by another format class object, use that. **/
            IF JSO_Num_Formats:HAS( ch_FormatCode ) THEN
                ob_Format:num_format_index = JSO_Num_Formats:GetInteger(ch_FormatCode ).
            ELSE
            DO:
                JSO_Num_Formats:ADD(ch_FormatCode, in_Index).
                ob_Format:num_format_index = in_Index. /*** Duck ***/
                
                CREATE ttNumFormats.

                ASSIGN
                    ttNumFormats.number_format_index = in_Index
                    ttNumFormats.format_code         = ch_FormatCode.

                in_Index = in_Index + 1.
            END.

            in_num_format_count = in_num_format_count + 1.

/*             ttxf_Formats.format_index = in_num_format_count. */
        
            THIS-OBJECT:num_format_count = in_num_format_count.
        END.

    END METHOD.

    METHOD PRIVATE VOID _prepare_borders():
    END METHOD.

    METHOD PRIVATE VOID _prepare_fills():
    END METHOD.

    METHOD PUBLIC VOID _assemble_xml_file():

        DEFINE VARIABLE ObjClassInstance     AS CLASS Progress.Lang.Object .
        DEFINE VARIABLE ObjWorkSheetInstance AS CLASS worksheet.
            
        MESSAGE '_prepare_format_properties()'.
        THIS-OBJECT:_prepare_format_properties().

        /** Create a new Worksheet if one has not been defined. **/
        IF NOT CAN-FIND(FIRST ttWorkSheet) THEN 
            THIS-OBJECT:add_worksheet().
        
        /** Write the workbook element. **/
        THIS-OBJECT:_write_workbook().

        /** Close the file. **/
        THIS-OBJECT:_xml_close().

    END METHOD.



    /** PRIVATE METHODS... **/

    METHOD PRIVATE VOID _store_workbook():
        
        DEFINE VARIABLE chXMLFiles      AS CHARACTER    NO-UNDO.
        DEFINE VARIABLE objPackager     AS CLASS packager NO-UNDO.
        DEFINE VARIABLE INETIME         AS INTEGER     NO-UNDO.
        DEFINE VARIABLE chOSCOMMAND     AS CHARACTER   NO-UNDO.

        ETIME(true).

        THIS-OBJECT:_prepare_sst_string_data().
        
        THIS-OBJECT:_prepare_vml().
        
        THIS-OBJECT:_prepare_defined_names().
        
        THIS-OBJECT:_prepare_drawings().
        
        THIS-OBJECT:_add_chart_data().
        
        THIS-OBJECT:_prepare_tables().
        
        MESSAGE 'NEW Package' etime.
        objPackager = NEW Packager().

        objPackager:_add_Workbook(THIS-OBJECT).
        objPackager:_set_tmpdir( THIS-OBJECT:TempDir ).

        MESSAGE '_create_package start' etime.
         chXMLFiles = objPackager:_create_package().
        MESSAGE '_create_package end' etime.

        MESSAGE etime.

        
/*         OS-DELETE VALUE(THIS-OBJECT:filename). */

        FILE-INFO:FILE-NAME = '.'.

        chOSCOMMAND = SUBSTITUTE('cd "&1"&&&&&4/zip -S -r &4/&2 &3', 
                                 objPackager:temp-dir,
                                 THIS-OBJECT:filename,
                                 chXMLFiles,
                                 FILE-INFO:FULL-PATHNAME).

        message chOSCOMMAND.
        MESSAGE etime.
        OS-COMMAND VALUE(chOSCOMMAND).
        MESSAGE etime.

        DELETE OBJECT objPackager.            
    END METHOD.

    METHOD PRIVATE VOID _write_workbook():
            

        hnSAXWriter:START-ELEMENT('workbook').

        hnSAXWriter:DECLARE-NAMESPACE('http://schemas.openxmlformats.org/spreadsheetml/2006/main').
        hnSAXWriter:DECLARE-NAMESPACE('http://schemas.openxmlformats.org/officeDocument/2006/relationships','r').

            /** Write the fileVersion element. **/
        THIS-OBJECT:_write_file_version().

        /** Write the workbookPr element. **/
        THIS-OBJECT:_write_workbook_pr().

        /** Write the bookViews element. **/
        THIS-OBJECT:_write_book_views().

        /** Write the sheets element. **/
        THIS-OBJECT:_write_sheets().

        /** Write the workbook defined names. **/
        THIS-OBJECT:_write_defined_names().

        /** Write the calcPr element. **/
        THIS-OBJECT:_write_calc_pr().
        
        hnSAXWriter:END-ELEMENT('workbook').

    END METHOD.

    METHOD PRIVATE VOID _write_file_version():
        hnSAXWriter:WRITE-EMPTY-ELEMENT('fileVersion').
        hnSAXWriter:INSERT-ATTRIBUTE('appName', 'xl').
        hnSAXWriter:INSERT-ATTRIBUTE('lastEdited', '5').
        hnSAXWriter:INSERT-ATTRIBUTE('lowestEdited', '5').
        hnSAXWriter:INSERT-ATTRIBUTE('rupBuild', '9303').
    END METHOD.

    METHOD PRIVATE VOID _write_workbook_pr():
        hnSAXWriter:WRITE-EMPTY-ELEMENT('workbookPr').
        hnSAXWriter:INSERT-ATTRIBUTE('defaultThemeVersion', '124226').
    END METHOD.

    METHOD PRIVATE VOID _write_book_views():

        hnSAXWriter:START-ELEMENT('bookViews').
        hnSAXWriter:WRITE-EMPTY-ELEMENT('workbookView').
        
        hnSAXWriter:INSERT-ATTRIBUTE('xWindow', '480').
        hnSAXWriter:INSERT-ATTRIBUTE('yWindow', '30').
        hnSAXWriter:INSERT-ATTRIBUTE('windowWidth', '18195').
        hnSAXWriter:INSERT-ATTRIBUTE('windowHeight', '12075').
        
        hnSAXWriter:END-ELEMENT('bookViews').
    END METHOD.

    METHOD PRIVATE VOID _write_sheets():

        hnSAXWriter:START-ELEMENT('sheets').
        
        FOR EACH ttWorkSheet
            BY ttWorkSheet.sheetID:
            THIS-OBJECT:_write_sheet(ttWorkSheet.sheetname,ttWorkSheet.sheetID,ttWorkSheet.hidden).
        END.

        hnSAXWriter:END-ELEMENT('sheets').

    END METHOD.

    METHOD PRIVATE VOID _write_sheet(INPUT pchName    AS CHARACTER,
                                     INPUT pinSheetID AS INTEGER,
                                     INPUT plgHidden  AS LOGICAL):

        hnSAXWriter:WRITE-EMPTY-ELEMENT('sheet').
        hnSAXWriter:INSERT-ATTRIBUTE('name', pchName).
        hnSAXWriter:INSERT-ATTRIBUTE('sheetId', STRING(pinSheetID)).

        IF plgHidden THEN
            hnSAXWriter:INSERT-ATTRIBUTE('state', 'hidden').

        hnSAXWriter:INSERT-ATTRIBUTE('r:id', SUBSTITUTE("rId&1",pinSheetID) ).

    END METHOD.

    METHOD PRIVATE VOID _write_calc_pr():

        hnSAXWriter:WRITE-EMPTY-ELEMENT('calcPr').
        hnSAXWriter:INSERT-ATTRIBUTE('calcId', '145621').

    END METHOD.

    METHOD PRIVATE VOID _write_defined_names():
    END METHOD.

    METHOD PRIVATE VOID _prepare_sst_string_data():
    END METHOD.

    METHOD PRIVATE VOID _prepare_vml():
    END METHOD.

    METHOD PRIVATE VOID _prepare_defined_names():
    END METHOD.

    METHOD PRIVATE VOID _prepare_drawings():
    END METHOD.

    METHOD PRIVATE VOID _add_chart_data():
    END METHOD.

    METHOD PRIVATE VOID _prepare_tables():
    END METHOD.
    
    

END CLASS.

