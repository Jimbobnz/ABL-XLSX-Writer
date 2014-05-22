USING Progress.Json.ObjectModel.*.

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

    METHOD PUBLIC CHARACTER xl_colour (INPUT pch_Colour AS CHARACTER):
        
        DEFINE VARIABLE ch_Colour_Hex_Code AS CHARACTER   NO-UNDO.

        CASE pch_Colour:
            WHEN 'black' THEN
                ch_Colour_Hex_Code = '#000000'.
            WHEN 'blue' THEN
                ch_Colour_Hex_Code = '#0000FF'.
            WHEN 'brown' THEN
                ch_Colour_Hex_Code = '#800000'.
            WHEN 'cyan' THEN
                ch_Colour_Hex_Code = '#00FFFF'.
           WHEN 'gray' THEN
                ch_Colour_Hex_Code = '#808080'.
           WHEN 'green' THEN
                ch_Colour_Hex_Code = '#008000'.
            WHEN 'lime' THEN
                ch_Colour_Hex_Code = '#00FF00'.
            WHEN 'magenta' THEN
                ch_Colour_Hex_Code = '#FF00FF'.
            WHEN 'navy' THEN
                ch_Colour_Hex_Code = '#000080'.
            WHEN 'orange' THEN
                ch_Colour_Hex_Code = '#FF6600'.
            WHEN 'pink' THEN
                ch_Colour_Hex_Code = '#FF00FF'.
            WHEN 'purple' THEN
                ch_Colour_Hex_Code = '#800080'.
            WHEN 'red' THEN
                ch_Colour_Hex_Code = '#FF0000'.
            WHEN 'silver' THEN
                ch_Colour_Hex_Code = '#C0C0C0'.
            WHEN 'white' THEN
                ch_Colour_Hex_Code = '#FFFFFF'.
            WHEN 'yellow' THEN
                ch_Colour_Hex_Code = '#FFFF00'.
            OTHERWISE
                ch_Colour_Hex_Code = pch_Colour.

        END CASE.

        IF THIS-OBJECT:_validate_hex_colour(INPUT ch_Colour_Hex_Code) THEN
            RETURN REPLACE(ch_Colour_Hex_Code,'#','FF').
        ELSE
            MESSAGE SUBSTITUTE('&1 is an invalid colour choice.',
                               QUOTER(ch_Colour_Hex_Code)
                               )
                VIEW-AS ALERT-BOX WARNING TITLE 'Invalid Colour'.    
            
        RETURN ''.  /** If we got to this part of the code then return blank. No Colour.**/
            
    END METHOD.

    METHOD PRIVATE LOGICAL _validate_hex_colour (INPUT ch_Colour_Hex_Code AS CHARACTER):
        DEFINE VARIABLE in_Loop       AS INTEGER    NO-UNDO.
        DEFINE VARIABLE in_ASCII_Code AS INTEGER    NO-UNDO.
        
        ch_Colour_Hex_Code = CAPS(ch_Colour_Hex_Code).

        IF NOT (ch_Colour_Hex_Code MATCHES '#......') THEN 
            RETURN FALSE.

        DO in_Loop = 1 TO LENGTH( ch_Colour_Hex_Code ):
        
            in_ASCII_Code  = ASC( SUBSTRING(ch_Colour_Hex_Code,in_Loop,1) ).
        
            IF (NOT in_ASCII_Code GE 65 AND     /** 65 = A **/
                NOT in_ASCII_Code LE 70) OR     /** 70 = F **/
               (NOT in_ASCII_Code GE 48 AND     /** 48 = 0 **/
                NOT in_ASCII_Code LE 57) THEN   /** 57 = 9 **/
                RETURN FALSE.
                    
        END.

        RETURN TRUE.

    END METHOD.

END CLASS.
