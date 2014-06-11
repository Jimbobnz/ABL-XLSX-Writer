USING ABL-XLSX-Writer.ABL-XLSX-Writer.writer.*.
CLASS ABL-XLSX-Writer.ABL-XLSX-Writer.writer.PasswordHasher INHERITS BitwiseOperation:

    METHOD CHARACTER _encode_password (INPUT ch_Password AS CHARACTER):

        DEFINE VARIABLE in_loop         AS INTEGER      NO-UNDO.
        DEFINE VARIABLE ch_char         AS CHARACTER    NO-UNDO.
        DEFINE VARIABLE in_char         AS INT64        NO-UNDO.
        DEFINE VARIABLE in_password     AS INT64        NO-UNDO.
        DEFINE VARIABLE in_low_15       AS INT64        NO-UNDO.
        DEFINE VARIABLE in_high_15      AS INT64        NO-UNDO.
    
        in_password = 0x0000.
    
        DO in_loop = 1 TO LENGTH(ch_Password):
    
            ch_char     = SUBSTRING(ch_Password,in_loop,1).
            in_char     = ShiftLeft(ASC(ch_char), in_loop ).
            in_low_15   = BitAND(in_char, 0x7FFF ).
            in_high_15  = BitAND(in_char, ShiftLeft(0x7FFF, 15) ).
            in_high_15  = ShiftRight(in_high_15, 15).
            in_char     = XOR ( in_low_15, in_high_15).
            
            in_password = XOR(in_password, in_char).
        END.
        
        in_password = XOR(in_password, LENGTH(ch_Password)). 
        in_password = XOR(in_password, 0xCE4B).    
    
        RETURN CAPS( intToHex( in_password ) ).
    END METHOD.

END CLASS.
