CLASS ABL-XLSX-Writer.ABL-XLSX-Writer.writer.BitwiseOperation ABSTRACT:

    /** AND **/
    METHOD INT64 BitAND (INPUT in_Operand_A AS INT64,
                              INPUT in_Operand_B AS INT64):  

       DEFINE VARIABLE in_cbit     AS INTEGER     NO-UNDO.
       DEFINE VARIABLE in_result   AS INT64     NO-UNDO.
    
        DO in_cbit = 1 TO 64:
            IF LOGICAL( GET-BITS( in_Operand_A, in_cbit, 1 ) ) AND
               LOGICAL( GET-BITS( in_Operand_B, in_cbit, 1 ) ) THEN
                PUT-BITS( in_result, in_cbit, 1 ) = 1.
        END.
    
        RETURN in_result.
    END METHOD. /* End of FUNCTION BinaryAND */

    /** OR **/
    METHOD INT64 BITOR (INPUT in_Operand_A  AS INT64,
                             INPUT in_Operand_B  AS INT64):
 
        DEFINE VARIABLE in_cbit    AS INTEGER NO-UNDO.
        DEFINE VARIABLE in_result  AS INTEGER NO-UNDO.
   
        DO in_cbit = 1 TO 64:
            IF LOGICAL( GET-BITS( in_Operand_A, in_cbit, 1 ) ) OR
               LOGICAL( GET-BITS( in_Operand_B, in_cbit, 1 ) ) THEN
                PUT-BITS( in_result, in_cbit, 1 ) = 1.
        END.
 
        RETURN in_result.
     
    END METHOD. /*End of METHOD BinaryXOR */

    /** XOR **/
    METHOD INT64 XOR (INPUT intOperand1 AS INT64,
                              INPUT intOperand2 AS INT64):

        DEFINE VARIABLE iByteLoop  AS INTEGER NO-UNDO.
        DEFINE VARIABLE iXOResult  AS INT64 NO-UNDO.
        DEFINE VARIABLE lFirstBit  AS LOGICAL NO-UNDO.
        DEFINE VARIABLE lSecondBit AS LOGICAL NO-UNDO.
    
        iXOResult = 0.
    
        /*spin through each byte of each char*/
        DO iByteLoop = 1 TO 64: /* as processing a double byte character */
            /*find state (true / false) of each integer's byte*/
            ASSIGN
                lFirstBit  = LOGICAL(GET-BITS(intOperand1,iByteLoop  ,1))
                lSecondBit = LOGICAL(GET-BITS(intOperand2,iByteLoop , 1)).
    
            /* XOR each bit*/
            IF (lFirstBit  AND NOT lSecondBit) OR
               (lSecondBit AND NOT lFirstBit) THEN
                iXOResult = iXOResult + EXP(2, iByteLoop - 1).
        END.         

        RETURN iXOResult.
    END METHOD. /*End of METHOD BinaryXOR */

    /** ShiftRight **/
    METHOD INT64 ShiftRight (INPUT IN_Operand_A AS INT64,
                                     INPUT in_Operand_B AS INTEGER):
        /** Bit Shift Left **/
        RETURN INT64( in_Operand_A * EXP(2,in_Operand_B)  ).
    END METHOD.

    /** ShiftLeft **/
    METHOD INT64 ShiftLeft (INPUT IN_Operand_A AS INT64,
                            INPUT in_Operand_B AS INTEGER):
        
        RETURN INT64( in_Operand_A * EXP(2,in_Operand_B)  ).
    END METHOD.                 

METHOD CHARACTER intToHex (INPUT iNumber AS INT64) :

    DEF VAR chex  AS CHAR NO-UNDO.
    DEF VAR rbyte AS RAW  NO-UNDO.
    
    DO WHILE iNumber > 0:
      PUT-BYTE( rbyte, 1 ) = iNumber MODULO 256.
      chex = SUBSTITUTE('&1&2',
                        HEX-ENCODE( rbyte ),
                        chex).
      iNumber = TRUNCATE( iNumber / 256, 0 ).
    END.
    
    RETURN chex.
END METHOD.


END CLASS.
