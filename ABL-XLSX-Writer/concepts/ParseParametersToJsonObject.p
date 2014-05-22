

PROCEDURE parameters:

    DEFINE INPUT PARAMETER pch_Params   AS CHARACTER.

    DEFINE VARIABLE in_Num_Entries      AS INTEGER      NO-UNDO.
    DEFINE VARIABLE in_Entry            AS INTEGER      NO-UNDO.

    DEFINE VARIABLE ch_Object_Entry     AS CHARACTER    NO-UNDO.
    DEFINE VARIABLE ch_Object_Property  AS CHARACTER    NO-UNDO.
    DEFINE VARIABLE ch_Object_Value     AS CHARACTER    NO-UNDO.

    DEFINE VARIABLE lg_temp_value       AS LOGICAL      NO-UNDO.
    DEFINE VARIABLE de_temp_value       AS DECIMAL      NO-UNDO.
    DEFINE VARIABLE in_temp_value       AS DECIMAL      NO-UNDO.
    DEFINE VARIABLE ch_temp_value       AS CHARACTER    NO-UNDO.

    in_Num_Entries = NUM-ENTRIES(pch_Params,',').

    OBJECT-LOOP:
    DO in_Entry = 1 TO in_Num_Entries:

        ch_Object_Entry = ENTRY(in_Entry, pch_Params, ',').

        IF NUM-ENTRIES(ch_Object_Entry,':') EQ 2 THEN
        DO:
            
            ASSIGN
                ch_Object_Property = ENTRY(1,ch_Object_Entry, ':')
                ch_Object_Property = TRIM(ch_Object_Property)
                ch_Object_Property = TRIM(ch_Object_Property, "'")
                ch_Object_Property = TRIM(ch_Object_Property, '"')
                .

            ASSIGN
                ch_Object_Value    = ENTRY(2,ch_Object_Entry, ':')
                ch_Object_Value    = TRIM(ch_Object_Value)
                .

            /** LOGICAL VALUE **/
            IF ch_Object_Value EQ 'false' OR
               ch_Object_Value EQ 'true'  THEN
            DO:
                
                lg_temp_value = LOGICAL(ch_Object_Value,'true/false') NO-ERROR.

                IF NOT ERROR-STATUS:ERROR THEN 
                DO:
                    MESSAGE ch_Object_Property lg_temp_value 'Logical'.
                    NEXT OBJECT-LOOP.
                END.
                    
            END.

            /** CHARACTER VALUE - Detect for a quoted string **/
            IF ch_Object_Value MATCHES '"*"' OR
               ch_Object_Value MATCHES "'*'" THEN
            DO:
                ASSIGN
                    ch_temp_value = TRIM(ch_Object_Value, "'")
                    ch_temp_value = TRIM(ch_Object_Value, '"').
                
                MESSAGE ch_Object_Property ch_temp_value 'Character'.
                NEXT OBJECT-LOOP.
            END.

            /** DECIMAL VALUE - Test for a decimal point**/
            IF ch_Object_Value MATCHES "*~~.*" THEN
            DO:
                de_temp_value = DECIMAL(ch_Object_Value) NO-ERROR.

                IF NOT ERROR-STATUS:ERROR THEN 
                DO:
                    MESSAGE ch_Object_Property de_temp_value 'Decimal'.
                    NEXT OBJECT-LOOP.
                END.
                    
            END.

            /** INTEGER VALUE **/
            in_temp_value = INTEGER(ch_Object_Value) NO-ERROR.

            IF NOT ERROR-STATUS:ERROR THEN 
            DO:
                MESSAGE ch_Object_Property in_temp_value 'Integer'.
                NEXT OBJECT-LOOP.
            END.
        END.
    END.
END PROCEDURE.

RUN parameters(INPUT '"test":"123.01"').
