
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

CLASS ABL-XLSX-Writer.ABL-XLSX-Writer.writer.Packager:

    DEFINE PRIVATE PROPERTY filenames        AS CHARACTER GET. SET.
    DEFINE PRIVATE PROPERTY Workbook         AS CLASS workbook GET. SET.
    DEFINE PRIVATE PROPERTY objContentTypes  AS CLASS ContentTypes GET. SET.
    DEFINE PRIVATE PROPERTY workbook_count   AS INTEGER GET. SET.
    DEFINE PRIVATE PROPERTY unique_base_dir  AS CHARACTER GET. SET.
    DEFINE PUBLIC PROPERTY temp-dir          AS CHARACTER GET. PRIVATE SET.
    
    CONSTRUCTOR Packager():
        THIS-OBJECT:unique_base_dir = HEX-ENCODE(GENERATE-UUID).
    END.

    DESTRUCTOR Packager():

        /** Delete both the files contained within the directory and the directory itself**/
        OS-DELETE VALUE( THIS-OBJECT:temp-dir ) RECURSIVE.
    END.

    METHOD PUBLIC VOID _add_Workbook(INPUT workbook AS CLASS workbook):

        THIS-OBJECT:Workbook = workbook.
        THIS-OBJECT:workbook_count = THIS-OBJECT:workbook_count + 1.

    END METHOD.

    METHOD PUBLIC VOID _set_tmpdir(INPUT pchTempDir AS CHARACTER):

        pchTempDir = REPLACE(pchTempDir,'~\','/').
        pchTempDir = pchTempDir + '/' + THIS-OBJECT:unique_base_dir.
        pchTempDir = REPLACE(pchTempDir,'//','/').

        THIS-OBJECT:temp-dir = pchTempDir.

        OS-CREATE-DIR VALUE( THIS-OBJECT:temp-dir ).

    END METHOD.

    METHOD PRIVATE CHARACTER _filename(INPUT pchXMLFileName AS CHARACTER ):

        DEFINE VARIABLE chFullPathName  AS CHARACTER   NO-UNDO.

        THIS-OBJECT:filenames = THIS-OBJECT:filenames + QUOTER(pchXMLFileName) + ' '. 

        chFullPathName = THIS-OBJECT:temp-dir + '/' + pchXMLFileName.
        
        RETURN chFullPathName.

    END METHOD.

    METHOD PUBLIC CHARACTER  _create_package():
        
        MESSAGE '_write_worksheet_files()'.
        THIS-OBJECT:_write_worksheet_files().
        
        MESSAGE '_write_worksheet_files()'.
        THIS-OBJECT:_write_chartsheet_files().

        MESSAGE '_write_workbook_file()'.
        THIS-OBJECT:_write_workbook_file().

        MESSAGE '_write_chart_files()'.
        THIS-OBJECT:_write_chart_files().

        THIS-OBJECT:_write_drawing_files().
        THIS-OBJECT:_write_vml_files().
        THIS-OBJECT:_write_comment_files().
        THIS-OBJECT:_write_table_files().
        THIS-OBJECT:_write_shared_strings_file().
        THIS-OBJECT:_write_app_file(). 
        THIS-OBJECT:_write_core_file().
        THIS-OBJECT:_write_content_types_file().
        THIS-OBJECT:_write_styles_file().
        THIS-OBJECT:_write_theme_file().
        THIS-OBJECT:_write_root_rels_file().
        THIS-OBJECT:_write_workbook_rels_file().
        THIS-OBJECT:_write_worksheet_rels_files().
        THIS-OBJECT:_write_chartsheet_rels_files().
        THIS-OBJECT:_write_drawing_rels_files().
        THIS-OBJECT:_add_image_files().
        THIS-OBJECT:_add_vba_project().

        /** Return quoted spaced string. **/
        RETURN THIS-OBJECT:filenames.
    END METHOD.

    METHOD PUBLIC VOID _write_worksheet_files():

        /** The big one **/
        DEFINE VARIABLE hnTTWorksheets AS HANDLE      NO-UNDO.
        DEFINE VARIABLE hnBuffer       AS HANDLE      NO-UNDO.
        DEFINE VARIABLE hnQuery        AS HANDLE      NO-UNDO.
        DEFINE VARIABLE ObjWorkSheet   AS CLASS worksheet   NO-UNDO.
        DEFINE VARIABLE obJsonArray    AS CLASS JsonArray.
        
        DEFINE VARIABLE WorkSheetSelected AS LOGICAL INITIAL FALSE     NO-UNDO.

        /** Get the temp-table handle ..**/
        hnTTWorksheets = THIS-OBJECT:Workbook:worksheets().

        CREATE QUERY hnQuery.
        hnBuffer = hnTTWorksheets:DEFAULT-BUFFER-HANDLE.
        hnQuery:ADD-BUFFER( hnBuffer ).
        DEFINE VARIABLE chQueryWhere AS CHARACTER   NO-UNDO.
        
        hnQuery:QUERY-PREPARE( 'FOR EACH ttWorkSheet'  ).
        hnQuery:QUERY-OPEN().

        IF hnQuery:IS-OPEN THEN
        DO:

            hnQuery:GET-FIRST() NO-ERROR.       
    
            REPEAT WHILE NOT hnQuery:QUERY-OFF-END:
                             
                /** Get the Class Object Handle.. **/
                ObjWorkSheet = hnBuffer::ObjWorkSheet.
                
                /** Select the first work sheet as Selective Tab **/
                IF NOT WorkSheetSelected THEN 
                DO:
                    ObjWorkSheet:selected = TRUE.
                    WorkSheetSelected     = TRUE.
                END.
    
                ObjWorkSheet:_set_xml_writer( THIS-OBJECT:_filename( SUBSTITUTE('xl/worksheets/sheet&1.xml', hnBuffer::SheetID )) ).
                ObjWorkSheet:_assemble_xml_file().
    
                hnQuery:GET-NEXT() NO-ERROR.
            END.
            hnQuery:QUERY-CLOSE().
        END.

        DELETE OBJECT hnQuery.

    END METHOD.

    METHOD PUBLIC VOID _write_chartsheet_files():
    END METHOD.

    METHOD PUBLIC VOID _write_workbook_file():
        
        THIS-OBJECT:Workbook:_set_xml_writer( THIS-OBJECT:_filename('xl/workbook.xml') ).

        MESSAGE 'Workbook:_assemble_xml_file()'.
        THIS-OBJECT:Workbook:_assemble_xml_file().
    END METHOD.

    METHOD PUBLIC VOID _write_chart_files():
    END METHOD.

    METHOD PUBLIC VOID _write_drawing_files():
    END METHOD.

    METHOD PUBLIC VOID _write_vml_files():
    END METHOD.

    METHOD PUBLIC VOID _write_comment_files():
    END METHOD.

    METHOD PUBLIC VOID _write_table_files():
    END METHOD.

    METHOD PUBLIC VOID _write_shared_strings_file():

        DEFINE VARIABLE objsst AS CLASS SharedStrings.

        objsst = NEW SharedStrings(). 
        objsst:SharedStringTable = THIS-OBJECT:Workbook:SharedStringTable. 
        
        IF NOT LOGICAL(THIS-OBJECT:Workbook:SharedStringTable:count) THEN 
            RETURN.

        objsst:_set_xml_writer( THIS-OBJECT:_filename('xl/sharedStrings.xml') ).
        objsst:_assemble_xml_file().

    END METHOD.

    METHOD PUBLIC VOID _write_app_file():

        DEFINE VARIABLE hnTTWorksheets AS HANDLE      NO-UNDO.
        DEFINE VARIABLE hnBuffer       AS HANDLE      NO-UNDO.
        DEFINE VARIABLE hnQuery        AS HANDLE      NO-UNDO.
        DEFINE VARIABLE objApp         AS CLASS App   NO-UNDO.
        DEFINE VARIABLE obJsonArray    AS CLASS JsonArray.

        objApp      = NEW App().
        obJsonArray = NEW JsonArray().

/*         obJsonArray:ADD('Worksheets').                         */
/*         obJsonArray:ADD( STRING(THIS-OBJECT:workbook_count) ). */
/*         objApp:_add_heading_pair(INPUT obJsonArray ).          */
        
        hnTTWorksheets = THIS-OBJECT:Workbook:worksheets().

        CREATE QUERY hnQuery.
        hnBuffer = hnTTWorksheets:DEFAULT-BUFFER-HANDLE.
        hnQuery:ADD-BUFFER( hnBuffer ).
        DEFINE VARIABLE chQueryWhere AS CHARACTER   NO-UNDO.
        chQueryWhere = SUBSTITUTE('PRESELECT EACH &1', hnTTWorksheets:name).
        hnQuery:QUERY-PREPARE( chQueryWhere  ).
        hnQuery:QUERY-OPEN().

        obJsonArray:ADD('Worksheets').
        obJsonArray:ADD( STRING( hnQuery:NUM-RESULTS ) ).
        objApp:_add_heading_pair(INPUT obJsonArray ).


        hnQuery:GET-FIRST() NO-ERROR.       
        REPEAT WHILE NOT hnQuery:QUERY-OFF-END:
             objApp:_add_part_name(INPUT STRING(hnBuffer::SheetName)). 
            hnQuery:GET-NEXT() NO-ERROR.
        END.
        hnQuery:QUERY-CLOSE().
        
        DELETE OBJECT hnQuery.

        objApp:_set_xml_writer( THIS-OBJECT:_filename('docProps/app.xml') ).
        objApp:_assemble_xml_file().

    END METHOD.

    METHOD PUBLIC VOID _write_core_file():
       /** Write the _rels/.rels xml file. **/
        DEFINE VARIABLE objCore         AS CLASS Core NO-UNDO.
        DEFINE VARIABLE hn_Properties   AS HANDLE      NO-UNDO.

        objCore = NEW Core().

        /** Get the handle of the temp-table "properties". **/
        hn_Properties = THIS-OBJECT:Workbook:properties().

        /** Pass the handle of temp-table to the core object.. **/
        objCore:_Set_Properties(INPUT hn_Properties).

        objCore:_set_xml_writer( THIS-OBJECT:_filename('docProps/core.xml') ).
        objCore:_assemble_xml_file().

        DELETE OBJECT objCore.

    END METHOD.

    METHOD PUBLIC VOID _write_content_types_file():

        THIS-OBJECT:objContentTypes = NEW ContentTypes().

        DEFINE VARIABLE hnTTWorksheets AS HANDLE      NO-UNDO.

        /** Return the temp-table handle for the worksheet.**/
        hnTTWorksheets = THIS-OBJECT:Workbook:worksheets().
        
        IF NOT hnTTWorksheets:HAS-RECORDS THEN
            THIS-OBJECT:objContentTypes:_add_worksheet_name(INPUT "Sheet1").
        ELSE
        DO:

            DEFINE VARIABLE hnQuery AS HANDLE      NO-UNDO.
            DEFINE VARIABLE hnBuffer AS HANDLE      NO-UNDO.

            CREATE QUERY hnQuery.
            hnBuffer = hnTTWorksheets:DEFAULT-BUFFER-HANDLE.
            hnQuery:ADD-BUFFER( hnBuffer ).
            DEFINE VARIABLE chQueryWhere AS CHARACTER   NO-UNDO.
            chQueryWhere = SUBSTITUTE('FOR EACH &1', hnTTWorksheets:name).
            hnQuery:QUERY-PREPARE( chQueryWhere  ).
            hnQuery:QUERY-OPEN().
            hnQuery:GET-FIRST() NO-ERROR.       
            REPEAT WHILE NOT hnQuery:QUERY-OFF-END:
                THIS-OBJECT:objContentTypes:_add_worksheet_name(INPUT 'sheet' + hnBuffer::SheetID).
                hnQuery:GET-NEXT() NO-ERROR.
            END.
            hnQuery:QUERY-CLOSE().

            DELETE OBJECT hnQuery.
        END.

        THIS-OBJECT:objContentTypes:_set_xml_writer( THIS-OBJECT:_filename('[Content_Types].xml') ).
        THIS-OBJECT:objContentTypes:_assemble_xml_file().

    END METHOD.

    METHOD PUBLIC VOID _write_styles_file():

        DEFINE VARIABLE objStyles AS CLASS Styles.

        objStyles = NEW Styles().

        /** Return the temp-table handle for the xf_formats objects.**/
        objStyles:_set_xf_formats( INPUT THIS-OBJECT:Workbook:xf_formats(),
                                   INPUT THIS-OBJECT:Workbook:fonts(),
                                   INPUT THIS-OBJECT:Workbook:Num_Format()
                                   ).


        /** Write the population stuff here....**/

        objStyles:_set_xml_writer(THIS-OBJECT:_filename('xl/styles.xml')).
        objStyles:_assemble_xml_file().

        DELETE OBJECT objStyles.
    END METHOD.

    METHOD PUBLIC VOID _write_theme_file():

        DEFINE VARIABLE objTheme AS CLASS Theme.

        objTheme = NEW Theme().

        objTheme:_set_xml_writer(THIS-OBJECT:_filename('xl/theme/theme1.xml')).
        objTheme:_assemble_xml_file().

        DELETE OBJECT objTheme.

    END METHOD.

    METHOD PUBLIC VOID _write_root_rels_file():

       /** Write the _rels/.rels xml file. **/
        DEFINE VARIABLE objRelationships AS CLASS Relationships NO-UNDO.

        objRelationships = NEW Relationships().
        
        objRelationships:_add_document_relationship('/officeDocument', 'xl/workbook.xml').
        objRelationships:_add_package_relationship('/metadata/core-properties', 'docProps/core.xml').
        objRelationships:_add_document_relationship('/extended-properties', 'docProps/app.xml').

        objRelationships:_set_xml_writer( THIS-OBJECT:_filename('_rels/.rels') ).
        objRelationships:_assemble_xml_file().

        DELETE OBJECT objRelationships.

    END METHOD.

    METHOD PUBLIC VOID _write_workbook_rels_file():

        DEFINE VARIABLE objRelationships AS CLASS Relationships NO-UNDO.
        objRelationships = NEW Relationships().

        DEFINE VARIABLE hnTTWorksheets  AS HANDLE      NO-UNDO.
        DEFINE VARIABLE hnQuery         AS HANDLE      NO-UNDO.
        DEFINE VARIABLE hnBuffer        AS HANDLE      NO-UNDO.

        /** Return the temp-table handle for the worksheet.**/
        hnTTWorksheets = THIS-OBJECT:Workbook:worksheets().

/*         worksheet_index = 1  */
/*         chartsheet_index = 1 */
        CREATE QUERY hnQuery.
        hnBuffer = hnTTWorksheets:DEFAULT-BUFFER-HANDLE.
        hnQuery:ADD-BUFFER( hnBuffer ).
        DEFINE VARIABLE chQueryWhere AS CHARACTER   NO-UNDO.
        chQueryWhere = SUBSTITUTE('FOR EACH &1', hnTTWorksheets:name).
        hnQuery:QUERY-PREPARE( chQueryWhere  ).
        hnQuery:QUERY-OPEN().
        hnQuery:GET-FIRST() NO-ERROR.       
        REPEAT WHILE NOT hnQuery:QUERY-OFF-END:
            objRelationships:_add_document_relationship(INPUT '/worksheet',
                                                        INPUT SUBSTITUTE('worksheets/sheet&1.xml',
                                                                         hnBuffer::SheetID)
                                                        ).
            hnQuery:GET-NEXT() NO-ERROR.
        END.
        hnQuery:QUERY-CLOSE().

        objRelationships:_add_document_relationship('/theme', 'theme/theme1.xml').
        objRelationships:_add_document_relationship('/styles', 'styles.xml').

        /* Add the sharedString rel if there is string data in the workbook. */
        IF THIS-OBJECT:workbook:SharedStringTable:_has_recods() THEN
            objRelationships:_add_document_relationship(INPUT '/sharedStrings',
                                                        INPUT 'sharedStrings.xml').

        objRelationships:_set_xml_writer( THIS-OBJECT:_filename('xl/_rels/workbook.xml.rels') ).
        objRelationships:_assemble_xml_file().
        DELETE OBJECT objRelationships.

    END METHOD.

    METHOD PUBLIC VOID _write_worksheet_rels_files():
    END METHOD.

    METHOD PUBLIC VOID _write_chartsheet_rels_files():
    END METHOD.
    
    METHOD PUBLIC VOID _write_drawing_rels_files():
    END METHOD.

    METHOD PUBLIC VOID _add_image_files():
    END METHOD.

    METHOD PUBLIC VOID _add_vba_project():
    END METHOD.




END CLASS.
