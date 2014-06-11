
/*                     public static function hashPassword($pPassword = '') {            */
/*         $password   = 0x0000;                                                         */
/*         $charPos    = 1;       // char position                                       */
/*                                                                                       */
/*         // split the plain text password in its component characters                  */
/*         $chars = preg_split('//', $pPassword, -1, PREG_SPLIT_NO_EMPTY);               */
/*         foreach ($chars as $char) {                                                   */
/*             $value          = ord($char) << $charPos++; // shifted ASCII value        */
/*             $rotated_bits   = $value >> 15;             // rotated bits beyond bit 15 */
/*             $value          &= 0x7fff;                  // first 15 bits              */
/*             $password       ^= ($value | $rotated_bits);                              */
/*         }                                                                             */
/*                                                                                       */
/*         $password ^= strlen($pPassword);                                              */
/*         $password ^= 0xCE4B;                                                          */
/*                                                                                       */
/*         return(strtoupper(dechex($password)));                                        */
/*     }                                                                                 */

                   
/* DEFINE VARIABLE in_EXP AS INT EXTENT 16 INITIAL [ 1, 2, 4, 8, 16, 32, 64, 128, 256, 512, 1024, 2048, 4096, 8192, 16384, 32768 ]. */
DEFINE VARIABLE in_EXP AS INT EXTENT 16 INITIAL [ 1, 2, 4, 8, 0x10, 0x20, 0x40, 0x80, 0x100, 0x200, 0x400, 0x800, 0x1000, 0x2000, 0x4000, 0x8000 ].

FUNCTION ShiftLeft RETURNS INT64(INPUT IN_Operand_A AS INT64,
                                 INPUT in_Operand_B AS INTEGER):
  /** Bit Shift Left **/
  RETURN INT64( in_Operand_A * EXP(2,in_Operand_B)  ).

END FUNCTION.
 
FUNCTION ShiftRight RETURNS INT64(INPUT in_Operand_A AS INT64,
                                  INPUT in_Operand_B AS INTEGER):
  /** Bit Shift Right **/
/*   RETURN INTEGER( TRUNCATE( in_Operand_A / in_EXP[ in_Operand_B + 1], 0 ) ). */

   RETURN INT64( TRUNCATE( in_Operand_A / EXP(2,in_Operand_B), 0 ) ).

END FUNCTION.

FUNCTION BinaryAND RETURNS INTEGER (INPUT in_Operand_A AS INT64,
                                    INPUT in_Operand_B AS INT64):  

   DEFINE VARIABLE in_cbit     AS INTEGER     NO-UNDO.
   DEFINE VARIABLE in_result   AS INT64     NO-UNDO.

   DO in_cbit = 1 TO 64:
      IF LOGICAL( GET-BITS( in_Operand_A, in_cbit, 1 ) ) AND
         LOGICAL( GET-BITS( in_Operand_B, in_cbit, 1 ) )
      THEN
         PUT-BITS( in_result, in_cbit, 1 ) = 1.
  END.

  RETURN in_result.

END FUNCTION. /* End of FUNCTION BinaryAND */

FUNCTION BinaryOR RETURNS INT64(INPUT in_Operand_A  AS INT64,
                                INPUT in_Operand_B  AS INT64):
 
      DEFINE VARIABLE in_cbit    AS INTEGER NO-UNDO.
      DEFINE VARIABLE in_result  AS INTEGER NO-UNDO.
   
      DO in_cbit = 1 TO 64:
         IF LOGICAL( GET-BITS( in_Operand_A, in_cbit, 1 ) ) OR
            LOGICAL( GET-BITS( in_Operand_B, in_cbit, 1 ) )
         THEN
             PUT-BITS( in_result, in_cbit, 1 ) = 1.
      END.
 
      RETURN in_result.
 
END FUNCTION. /*End of METHOD BinaryXOR */

FUNCTION BinaryXOR RETURNS INT64(INPUT intOperand1 AS INT64,
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
END FUNCTION. /*End of METHOD BinaryXOR */

FUNCTION dec2hex RETURNS character ( INPUT iNumber AS INT64 ) :
 
    DEFINE VARIABLE in_Rem      AS INTEGER      NO-UNDO.
    DEFINE VARIABLE ch_HEX      AS CHARACTER    NO-UNDO.
    DEFINE VARIABLE ch_hexCode  AS CHARACTER    NO-UNDO EXTENT 16 INITIAL ['0','1','2','3','4','5','6','7','8','9','a','b','c','d','e','f'].
 
    IF NOT LOGICAL(iNumber) THEN 
        RETURN "0".
    ELSE 
        IF iNumber LT 0 THEN 
            RETURN ERROR.
 
    DO WHILE LOGICAL(iNumber):
        ASSIGN
            in_Rem  = iNumber MODULO 16
            iNumber = TRUNCATE(iNumber / 16,0)
            ch_HEX = SUBSTITUTE('&1&2',                           
                                 ch_hexCode[in_Rem + 1],
                                 ch_HEX).
    END.

    RETURN ch_HEX.
END FUNCTION.
 
FUNCTION _encode_password RETURNS CHARACTER(INPUT ch_Password AS CHARACTER):

    DEFINE VARIABLE in_loop         AS INTEGER      NO-UNDO.
    DEFINE VARIABLE ch_char         AS CHARACTER    NO-UNDO.
    DEFINE VARIABLE in_char         AS INT64        NO-UNDO.
    DEFINE VARIABLE in_password     AS INT64        NO-UNDO.
    DEFINE VARIABLE in_low_15       AS INT64        NO-UNDO.
    DEFINE VARIABLE in_high_15      AS INT64        NO-UNDO.

/*     IF LENGTH(ch_Password) GT 15 THEN                                  */
/*         RETURN ERROR "Maximum Password length is 15 characters long.". */
        
    in_password = 0x0000.

    DO in_loop = 1 TO LENGTH(ch_Password):

        ch_char     = SUBSTRING(ch_Password,in_loop,1).
        in_char     = ShiftLeft(ASC(ch_char), in_loop ).
        in_low_15   = BinaryAND(in_char, 0x7FFF ).
        in_high_15  = BinaryAND(in_char, ShiftLeft(0x7FFF, 15) ).
        in_high_15  = ShiftRight(in_high_15, 15).
        in_char     = BinaryXOR ( in_low_15, in_high_15).
        
        in_password = BinaryXOR(in_password, in_char).
    END.
    
    in_password = BinaryXOR(in_password, LENGTH(ch_Password)). 
    in_password = BinaryXOR(in_password, 0xCE4B).    
    
    RETURN CAPS( dec2hex(in_password) ).

END FUNCTION.


message _encode_password('test') .

IF error-status:error  THEN
    MESSAGE  RETURN-VALUE
        VIEW-AS ALERT-BOX ERROR.

