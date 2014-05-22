
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

CLASS ABL-XLSX-Writer.ABL-XLSX-Writer.writer.Core INHERITS xmlwriter:
        
    DEFINE VARIABLE obj_Core_Properties AS CLASS JsonObject NO-UNDO.
    DEFINE VARIABLE data                AS CHARACTER        NO-UNDO.

    DEFINE PRIVATE PROPERTY localtime AS DATETIME-TZ NO-UNDO GET. SET.
    
    CONSTRUCTOR Core():
        THIS-OBJECT:obj_Core_Properties = NEW JsonObject().
        THIS-OBJECT:localtime           = NOW.
    END CONSTRUCTOR.

    DESTRUCTOR Core():
        
        /** House keeping.**/
        IF VALID-OBJECT(obj_Core_Properties) THEN
            DELETE OBJECT obj_Core_Properties.

    END DESTRUCTOR.
    
    METHOD PUBLIC VOID _assemble_xml_file():

        THIS-OBJECT:_write_cp_core_properties().

        /** Close the file.**/
        THIS-OBJECT:_xml_close().

    END METHOD.

    METHOD PUBLIC VOID _set_properties(INPUT hn_Properties AS HANDLE):
        
        DEFINE VARIABLE hnPropertyBuffer AS HANDLE      NO-UNDO.

        /* Validation check*/
        IF NOT VALID-HANDLE(hn_Properties)      OR 
           hn_Properties:TYPE NE 'TEMP-TABLE' THEN
            RETURN.
        
        hnPropertyBuffer = hn_Properties:DEFAULT-BUFFER-HANDLE.

        hnPropertyBuffer:FIND-FIRST() NO-ERROR.

        /** For whatever reason that there is no property record, move on.**/
        IF NOT hnPropertyBuffer:AVAILABLE THEN
            RETURN.

        obj_Core_Properties:READ(INPUT  hnPropertyBuffer ).
        
    END METHOD.

    
    METHOD PRIVATE VOID _write_cp_core_properties():
        
        hnSAXWriter:START-ELEMENT('cp:coreProperties').

    hnSAXWriter:DECLARE-NAMESPACE('http://schemas.openxmlformats.org/package/2006/metadata/core-properties','cp').
    hnSAXWriter:DECLARE-NAMESPACE('http://purl.org/dc/elements/1.1/','dc').
    hnSAXWriter:DECLARE-NAMESPACE('http://purl.org/dc/terms/','dcterms').
    hnSAXWriter:DECLARE-NAMESPACE('http://purl.org/dc/dcmitype/','dcmitype').
    hnSAXWriter:DECLARE-NAMESPACE('http://www.w3.org/2001/XMLSchema-instance','xsi').

        THIS-OBJECT:_write_dc_creator().
        THIS-OBJECT:_write_cp_last_modified_by().
        THIS-OBJECT:_write_dcterms_created().
        THIS-OBJECT:_write_dcterms_modified().
        THIS-OBJECT:_write_dc_title().
        THIS-OBJECT:_write_dc_subject().
        THIS-OBJECT:_write_cp_keywords().
        THIS-OBJECT:_write_dc_description().
        THIS-OBJECT:_write_cp_category().
        THIS-OBJECT:_write_cp_content_status().
        
        hnSAXWriter:END-ELEMENT('cp:coreProperties').

    END METHOD.

    METHOD PRIVATE VOID _write_dc_creator():

        /** Manditory Field**/
        IF obj_Core_Properties:HAS('author') THEN
            data = obj_Core_Properties:GetCharacter('author').
        ELSE
            data = 'ABL xslx writer on OpenEdge ' + PROVERSION(1). 
        
        hnSAXWriter:WRITE-DATA-ELEMENT('dc:creator',data).

    END METHOD.

    METHOD PRIVATE VOID _write_cp_last_modified_by():

        /** Manditory Field**/
        IF obj_Core_Properties:HAS('author') THEN
            data = obj_Core_Properties:GetCharacter('author').
        ELSE
            data =  'ABL xslx writer on OpenEdge ' + PROVERSION(1).
        
        hnSAXWriter:WRITE-DATA-ELEMENT('cp:lastModifiedBy',data).
        

    END METHOD.

    METHOD PRIVATE VOID _write_dcterms_created():

        hnSAXWriter:START-ELEMENT('dcterms:created').
        hnSAXWriter:INSERT-ATTRIBUTE('xsi:type', 'dcterms:W3CDTF').

        DEFINE VARIABLE dtz-LocalTime AS DATETIME-TZ        NO-UNDO.

        
        dtz-LocalTime = DATETIME-TZ( THIS-OBJECT:localtime, 0 ). /** O = UTC/GMT Little bit crude**/
        hnSAXWriter:WRITE-CHARACTERS(SUBSTITUTE("&1T&2Z",
                                                ISO-DATE(DATE(dtz-LocalTime)),
                                                STRING(INTEGER(MTIME(dtz-LocalTime) / 1000),"HH:MM:SS")
                                                )
                                       ).
        hnSAXWriter:END-ELEMENT('dcterms:created').


    END METHOD.

    METHOD PRIVATE VOID _write_dcterms_modified():
        

        hnSAXWriter:START-ELEMENT('dcterms:modified').
        hnSAXWriter:INSERT-ATTRIBUTE('xsi:type', 'dcterms:W3CDTF').

        DEFINE VARIABLE dtz-LocalTime AS DATETIME-TZ        NO-UNDO.
        dtz-LocalTime = DATETIME-TZ( THIS-OBJECT:localtime, 0 ). /** O = UTC/GMT Little bit crude**/
        hnSAXWriter:WRITE-CHARACTERS(SUBSTITUTE("&1T&2Z",
                                                ISO-DATE(DATE(dtz-LocalTime)),
                                                STRING(INTEGER(MTIME(dtz-LocalTime) / 1000),"HH:MM:SS")
                                                )
                                       ).
        hnSAXWriter:END-ELEMENT('dcterms:modified').

    END METHOD.

    METHOD PRIVATE VOID _write_dc_title():
        
        IF obj_Core_Properties:HAS('title') AND NOT obj_Core_Properties:IsNull('title') THEN
            hnSAXWriter:WRITE-DATA-ELEMENT('dc:title', obj_Core_Properties:GetCharacter('title')).

    END METHOD.

    METHOD PRIVATE VOID _write_dc_subject():

        IF obj_Core_Properties:HAS('subject') AND NOT obj_Core_Properties:IsNull('subject') THEN
            hnSAXWriter:WRITE-DATA-ELEMENT('dc:subject', obj_Core_Properties:GetCharacter('subject')).

    END METHOD.

    METHOD PRIVATE VOID _write_cp_keywords():

        IF obj_Core_Properties:HAS('keywords') AND NOT obj_Core_Properties:IsNull('keywords') THEN
            hnSAXWriter:WRITE-DATA-ELEMENT('cp:keywords', obj_Core_Properties:GetCharacter('keywords')).

    END METHOD.

    METHOD PRIVATE VOID _write_dc_description():
        
        IF obj_Core_Properties:HAS('comments')  AND NOT obj_Core_Properties:IsNull('comments') THEN
            hnSAXWriter:WRITE-DATA-ELEMENT('dc:description', obj_Core_Properties:GetCharacter('comments')).

    END METHOD.

    METHOD PRIVATE VOID _write_cp_category():

        IF obj_Core_Properties:HAS('category') AND NOT obj_Core_Properties:IsNull('category') THEN
            hnSAXWriter:WRITE-DATA-ELEMENT('cp:category', obj_Core_Properties:GetCharacter('category')).

    END METHOD.

    METHOD PRIVATE VOID _write_cp_content_status():

        IF obj_Core_Properties:HAS('status') AND NOT obj_Core_Properties:IsNull('status') THEN
            hnSAXWriter:WRITE-DATA-ELEMENT('cp:contentStatus', obj_Core_Properties:GetCharacter('status')).

    END METHOD.


END CLASS.


