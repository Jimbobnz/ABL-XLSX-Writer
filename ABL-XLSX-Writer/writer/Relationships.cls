
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

CLASS ABL-XLSX-Writer.ABL-XLSX-Writer.writer.Relationships INHERITS xmlwriter:
        
    DEFINE PRIVATE PROPERTY id AS INTEGER NO-UNDO GET. SET.

    DEFINE PRIVATE TEMP-TABLE ttRelationship NO-UNDO
        FIELD id    AS CHARACTER
        FIELD type  AS CHARACTER
        FIELD target AS CHARACTER.

    METHOD PUBLIC VOID _assemble_xml_file():

        THIS-OBJECT:_write_relationships().

        /** Close the file.**/
        THIS-OBJECT:_xml_close().

    END METHOD.

    METHOD PUBLIC VOID _add_document_relationship(INPUT rel_type AS CHARACTER,
                                                  INPUT target    AS CHARACTER):
        
        rel_type = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships' + rel_type.

        CREATE ttRelationship.

        ASSIGN
            ttRelationship.type  = rel_type 
            ttRelationship.target = target.

    END METHOD.

    METHOD PUBLIC VOID _add_package_relationship(INPUT rel_type AS CHARACTER,
                                                 INPUT target   AS CHARACTER):

        rel_type = 'http://schemas.openxmlformats.org/package/2006/relationships' + rel_type.

        CREATE ttRelationship.

        ASSIGN
            ttRelationship.type  = rel_type 
            ttRelationship.target = target.

    END METHOD.

    METHOD PUBLIC VOID _add_ms_package_relationship():

    END METHOD.

    METHOD PUBLIC VOID _add_worksheet_relationship(INPUT rel_type AS CHARACTER,
                                                   INPUT target   AS CHARACTER):

        rel_type = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships' + rel_type.

        CREATE ttRelationship.

        ASSIGN
            ttRelationship.type   = rel_type 
            ttRelationship.target = target.

    END METHOD.

    METHOD PRIVATE VOID _write_relationships():
        
        hnSAXWriter:START-ELEMENT('Relationships').

        hnSAXWriter:DECLARE-NAMESPACE('http://schemas.openxmlformats.org/package/2006/relationships').
        
        THIS-OBJECT:id = 0.

        FOR EACH ttRelationship:
            THIS-OBJECT:_write_relationship(INPUT ttRelationship.type,  
                                            INPUT ttRelationship.target).
        END.

        hnSAXWriter:END-ELEMENT('Relationships').

    END METHOD.

    METHOD PRIVATE VOID _write_relationship(INPUT rel_type AS CHARACTER,
                                            INPUT target    AS CHARACTER ):

        THIS-OBJECT:id = THIS-OBJECT:id + 1.

        hnSAXWriter:WRITE-EMPTY-ELEMENT('Relationship').
        hnSAXWriter:INSERT-ATTRIBUTE('Id', SUBSTITUTE('rId&1',THIS-OBJECT:id) ).
        hnSAXWriter:INSERT-ATTRIBUTE('Type', rel_type ).
        hnSAXWriter:INSERT-ATTRIBUTE('Target', target ).

    END METHOD.

END CLASS.
