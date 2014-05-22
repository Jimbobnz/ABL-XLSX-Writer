
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


CLASS ABL-XLSX-Writer.ABL-XLSX-Writer.writer.SharedStringTable:

    DEFINE TEMP-TABLE ttSharedStrings
        FIELD SharedString AS CHARACTER CASE-SENSITIVE
        FIELD IndexPos     AS INTEGER
        INDEX idxSharedString IS PRIMARY UNIQUE
            SharedString
        INDEX idxIndexPos IS UNIQUE
            IndexPos.

    DEFINE PUBLIC PROPERTY count        AS  INTEGER INITIAL 0 NO-UNDO GET. PRIVATE SET.
    DEFINE PUBLIC PROPERTY unique_count AS  INTEGER INITIAL 0 NO-UNDO GET. PRIVATE SET.
    
    METHOD PUBLIC INTEGER _get_shared_string_index (INPUT pchString AS CHARACTER):

        FIND ttSharedStrings
            WHERE ttSharedStrings.SharedString EQ pchString
            NO-ERROR.

        IF NOT AVAILABLE ttSharedStrings THEN
        DO:
            
            CREATE ttSharedStrings.

            ASSIGN
                ttSharedStrings.SharedString = pchString
                ttSharedStrings.IndexPos     = THIS-OBJECT:unique_count.

            THIS-OBJECT:unique_count = THIS-OBJECT:unique_count + 1.
        END.

        /* Count how many time this shared string is used accross all work sheets.*/    
        THIS-OBJECT:count        = THIS-OBJECT:count        + 1.

        RETURN ttSharedStrings.IndexPos.

    END METHOD.

    METHOD PUBLIC CHARACTER _get_shared_string(INPUT pinIndexPos AS INTEGER):

        FIND ttSharedStrings
            WHERE ttSharedStrings.IndexPos EQ pinIndexPos
            NO-ERROR.

        IF AVAILABLE ttSharedStrings THEN
            RETURN ttSharedStrings.SharedString.
        ELSE
            RETURN "".

    END METHOD.

    METHOD PUBLIC HANDLE _get_strings():
        RETURN TEMP-TABLE ttSharedStrings:HANDLE.
    END METHOD.

    METHOD PUBLIC LOGICAL _has_recods():
        RETURN TEMP-TABLE ttSharedStrings:HAS-RECORDS.
    END METHOD.


END CLASS.
