
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

CLASS ABL-XLSX-Writer.ABL-XLSX-Writer.writer.Utility:

    METHOD PUBLIC CHARACTER xl_col_to_name (INPUT col_num AS INTEGER):

        DEFINE VARIABLE remainder  AS INTEGER     NO-UNDO.
        DEFINE VARIABLE col_letter AS CHARACTER   NO-UNDO.
        DEFINE VARIABLE col_str    AS CHARACTER   NO-UNDO.

        DO WHILE col_num GT 0:
            remainder = (col_num - 1) MOD 26.
            col_letter = CHR( ASC('A') + remainder ).
            col_str = col_letter + col_str.
            col_num = INTEGER((col_num - remainder) / 26).
        END.

        RETURN col_str.

    END METHOD.

    METHOD PUBLIC CHARACTER xl_col_to_name (INPUT col_num AS INTEGER,
                                            INPUT Col_abs AS LOGICAL):

        DEFINE VARIABLE col_Name AS CHARACTER   NO-UNDO.

        col_Name = SUBSTITUTE("&1&2",
                              (IF col_abs THEN CHR(36) ELSE ''),
                              THIS-OBJECT:xl_col_to_name(INPUT col_num)
                              ).

        RETURN col_Name.
    END METHOD.
    
    METHOD PUBLIC CHARACTER xl_rowcol_to_cell (INPUT inRow   AS INTEGER,
                                               INPUT inCol AS INTEGER):

        DEFINE VARIABLE cell_Name AS CHARACTER   NO-UNDO.

        cell_Name = SUBSTITUTE("&1&2",
                               THIS-OBJECT:xl_col_to_name(INPUT inCol),
                               inRow).

        RETURN cell_Name.
    END METHOD.


    METHOD PUBLIC CHARACTER xl_rowcol_to_cell (INPUT inCustNum AS INTEGER):
      
        RETURN "".
    END METHOD.

    METHOD PUBLIC CHARACTER xl_rowcol_to_cell ():
RETURN "".
    END METHOD.

    

    METHOD PUBLIC CHARACTER xl_rowcol_to_cell (INPUT inRow AS INTEGER,
                                               INPUT iNCOL AS INTEGER,
                                               INPUT row_abs AS LOGICAL,
                                               INPUT Col_abs AS LOGICAL):


        DEFINE VARIABLE cell_Name AS CHARACTER   NO-UNDO.
        
        cell_Name = SUBSTITUTE("&1&2&3",
                               THIS-OBJECT:xl_col_to_name(INPUT inCol,
                                                          INPUT Col_abs),
                               (IF row_abs THEN '$' ELSE ''),
                               inRow).

        RETURN cell_Name.
    END METHOD.

END CLASS.
