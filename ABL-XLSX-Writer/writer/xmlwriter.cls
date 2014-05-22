
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

CLASS ABL-XLSX-Writer.ABL-XLSX-Writer.writer.xmlwriter:
    
    DEFINE PUBLIC PROPERTY hnSAXWriter AS HANDLE  GET. SET.           
    
    METHOD PUBLIC VOID _xml_declaration():
            
     IF NOT VALID-HANDLE(hnSAXWriter) THEN
         RETURN.

        THIS-OBJECT:hnSAXWriter:VERSION    = "1.0".
        THIS-OBJECT:hnSAXWriter:ENCODING   = "UTF-8".
        THIS-OBJECT:hnSAXWriter:STANDALONE = TRUE.

        THIS-OBJECT:hnSAXWriter:FORMATTED  = FALSE.
        THIS-OBJECT:hnSAXWriter:STRICT     = FALSE.

    END METHOD.

    METHOD VOID _set_xml_writer(INPUT pchXMLFile AS CHARACTER):

        IF VALID-HANDLE(hnSAXWriter) THEN
            THIS-OBJECT:hnSAXWriter:RESET().
        ELSE
            CREATE SAX-WRITER THIS-OBJECT:hnSAXWriter.

        DEFINE VARIABLE chDirectory AS CHARACTER   NO-UNDO.
        DEFINE VARIABLE i AS INTEGER     NO-UNDO.
        DO i = 1 TO (NUM-ENTRIES(pchXMLFile,'/') - 1):

            chDirectory = chDirectory + ENTRY(i, pchXMLFile,'/').

            IF chDirectory EQ '.' OR 
               chDirectory EQ '..' OR
               chDirectory EQ '' THEN
                NEXT.

            OS-CREATE-DIR VALUE(chDirectory).

            chDirectory = chDirectory + '/'.

        END.
        
        /** Write the XML declaration. **/
        THIS-OBJECT:_xml_declaration().
            
        
        THIS-OBJECT:hnSAXWriter:SET-OUTPUT-DESTINATION("file",pchXMLFile).
        THIS-OBJECT:hnSAXWriter:START-DOCUMENT( ).

    END METHOD.

    METHOD VOID _xml_close():
        
        IF NOT VALID-HANDLE(hnSAXWriter) THEN
            RETURN.

        THIS-OBJECT:hnSAXWriter:END-DOCUMENT( ).
        DELETE OBJECT THIS-OBJECT:hnSAXWriter.
        
    END METHOD.

END CLASS.
