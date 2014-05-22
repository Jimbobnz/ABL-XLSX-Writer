
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

CLASS ABL-XLSX-Writer.ABL-XLSX-Writer.writer.SharedStrings INHERITS xmlwriter:
    

    /** SharedStringTable is Set by the package class object **/
    DEFINE PUBLIC PROPERTY SharedStringTable  AS CLASS SharedStringTable GET. SET.
    
    METHOD PUBLIC VOID _assemble_xml_file():
            
        THIS-OBJECT:_write_sst().
        THIS-OBJECT:_xml_close().

    END METHOD.
    
    METHOD PRIVATE VOID _write_sst():

        hnSAXWriter:START-ELEMENT('sst').
        hnSAXWriter:DECLARE-NAMESPACE('http://schemas.openxmlformats.org/spreadsheetml/2006/main').

        hnSAXWriter:INSERT-ATTRIBUTE('count', STRING(THIS-OBJECT:SharedStringTable:count) ).
        hnSAXWriter:INSERT-ATTRIBUTE('uniqueCount', STRING(THIS-OBJECT:SharedStringTable:unique_count) ).

        THIS-OBJECT:_write_sst_strings().

        hnSAXWriter:END-ELEMENT('sst').

        THIS-OBJECT:_xml_close().

    END METHOD.

    METHOD PRIVATE VOID _write_sst_strings():

        /** Returns a temp-table handle **/
        DEFINE VARIABLE hnSST        AS HANDLE      NO-UNDO.
        DEFINE VARIABLE hnSSTBuffer  AS HANDLE      NO-UNDO.
        DEFINE VARIABLE hnSSTQuery   AS HANDLE      NO-UNDO.

        hnSST = THIS-OBJECT:SharedStringTable:_get_strings().

        IF NOT VALID-HANDLE(hnSST) THEN
            RETURN.
        
        hnSSTBuffer = hnSST:DEFAULT-BUFFER-HANDLE.

        CREATE QUERY hnSSTQuery.
        hnSSTQuery:ADD-BUFFER(hnSSTBuffer).
        hnSSTQuery:QUERY-PREPARE( 'FOR EACH ttSharedStrings BY ttSharedStrings.IndexPos' ).
        hnSSTQuery:QUERY-OPEN().
        hnSSTQuery:GET-FIRST() NO-ERROR.

        REPEAT WHILE NOT hnSSTQuery:QUERY-OFF-END:
            THIS-OBJECT:_write_si(INPUT hnSSTBuffer::SharedString).
            hnSSTQuery:GET-NEXT() NO-ERROR.
        END.

        hnSSTQuery:QUERY-CLOSE().

        DELETE OBJECT hnSSTQuery.

    END METHOD.

    METHOD PRIVATE VOID _write_si(INPUT chString AS CHARACTER):
        hnSAXWriter:START-ELEMENT('si').

        hnSAXWriter:START-ELEMENT( 't').
        
        /* Add attribute to preserve leading or trailing whitespace. */
        IF chString MATCHES " *" OR 
           chString MATCHES "* " THEN 
           hnSAXWriter:INSERT-ATTRIBUTE('xml:space', 'preserve').

        hnSAXWriter:WRITE-CHARACTERS( chString ). 
        hnSAXWriter:END-ELEMENT('t').

        hnSAXWriter:END-ELEMENT('si').
    END METHOD.

    

END CLASS.
