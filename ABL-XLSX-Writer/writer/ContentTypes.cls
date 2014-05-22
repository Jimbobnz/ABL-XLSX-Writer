
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

CLASS ABL-XLSX-Writer.ABL-XLSX-Writer.writer.ContentTypes INHERITS xmlwriter:

    DEFINE PRIVATE TEMP-TABLE ttDefault NO-UNDO
        FIELD Extension   AS CHARACTER 
        FIELD ContentType AS CHARACTER. 

    DEFINE PRIVATE TEMP-TABLE ttOveride NO-UNDO
        FIELD ContentType   AS CHARACTER 
        FIELD PartName      AS CHARACTER. 

    CONSTRUCTOR ContentTypes():

        CREATE ttDefault.
        ASSIGN
            ttDefault.Extension   = 'rels'
            ttDefault.ContentType = 'application/vnd.openxmlformats-package.relationships+xml'.

        CREATE ttDefault.
        ASSIGN
            ttDefault.Extension   = 'xml'
            ttDefault.ContentType = 'application/xml'.
            
        
        THIS-OBJECT:_add_override(INPUT '/docProps/app.xml',
                                  INPUT 'application/vnd.openxmlformats-officedocument.extended-properties+xml').

        THIS-OBJECT:_add_override(INPUT '/docProps/core.xml',
                                  INPUT 'application/vnd.openxmlformats-package.core-properties+xml').

        THIS-OBJECT:_add_override(INPUT '/xl/styles.xml',
                                  INPUT 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml').

        THIS-OBJECT:_add_override(INPUT '/xl/theme/theme1.xml',
                                  INPUT 'application/vnd.openxmlformats-officedocument.theme+xml').

        THIS-OBJECT:_add_override(INPUT '/xl/workbook.xml',
                                  INPUT 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml').

        THIS-OBJECT:_add_override(INPUT '/xl/sharedStrings.xml',
                                  INPUT 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml').

    END CONSTRUCTOR.





    METHOD VOID _assemble_xml_file():

        /** Write the workbook element. **/
        THIS-OBJECT:_write_types().

            /** Close the file. **/
        THIS-OBJECT:_xml_close().

    END METHOD.

    METHOD VOID _add_override(INPUT pchPartName    AS CHARACTER, 
                              INPUT pchContentType AS CHARACTER):

        CREATE ttOveride.
        ASSIGN 
            ttOveride.PartName    = pchPartName
            ttOveride.ContentType = pchContentType.

    END METHOD.

    METHOD VOID _add_worksheet_name(INPUT worksheet_name AS CHARACTER):

        worksheet_name = SUBSTITUTE('/xl/worksheets/&1.xml',
                                    worksheet_name).

        THIS-OBJECT:_add_override(INPUT worksheet_name,
                                  INPUT 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml').

    END METHOD.

    METHOD VOID _write_types():
    
        hnSAXWriter:START-ELEMENT('Types').

        hnSAXWriter:DECLARE-NAMESPACE('http://schemas.openxmlformats.org/package/2006/content-types').
        
        THIS-OBJECT:_write_defaults().
        THIS-OBJECT:_write_overrides().

        hnSAXWriter:END-ELEMENT('Types').

    END METHOD.

    METHOD VOID _write_defaults():

        FOR EACH TTDEFAULT:

            hnSAXWriter:WRITE-EMPTY-ELEMENT('Default').
            hnSAXWriter:INSERT-ATTRIBUTE('Extension',TTDEFAULT.Extension).
            hnSAXWriter:INSERT-ATTRIBUTE('ContentType',TTDEFAULT.ContentType).

        END.

    END METHOD.

    METHOD VOID _write_overrides():

        FOR EACH ttOveride:
    
            hnSAXWriter:WRITE-EMPTY-ELEMENT('Override').
            hnSAXWriter:INSERT-ATTRIBUTE('PartName', ttOveride.partName).
            hnSAXWriter:INSERT-ATTRIBUTE('ContentType',ttOveride.ContentType).
    
        END.
    
    END METHOD.

END CLASS.
