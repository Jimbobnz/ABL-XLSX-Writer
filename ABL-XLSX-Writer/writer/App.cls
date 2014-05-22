
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
USING Progress.Json.ObjectModel.*.

CLASS ABL-XLSX-Writer.ABL-XLSX-Writer.writer.App INHERITS xmlwriter:
        
    DEFINE VARIABLE myAppProp     AS CLASS JsonObject NO-UNDO.
    DEFINE VARIABLE heading_pairs AS CLASS JsonArray NO-UNDO.
    DEFINE VARIABLE part_name     AS CLASS JsonArray NO-UNDO.

    DEFINE VARIABLE data       AS CHARACTER   NO-UNDO.
    
    CONSTRUCTOR App():
        myAppProp     = NEW JsonObject().
        heading_pairs = NEW JsonArray().
        part_name     = NEW JsonArray().

    END CONSTRUCTOR.

    METHOD PUBLIC VOID _add_part_name (INPUT chPartName AS CHARACTER):

        DEFINE VARIABLE tempJsonArray AS CLASS JsonArray .

        tempJsonArray = NEW JsonArray().
        tempJsonArray:ADD( 'lpstr' ).
        tempJsonArray:ADD( chPartName ).

        THIS-OBJECT:part_name:ADD( tempJsonArray ).
    END METHOD.

    METHOD PUBLIC VOID _add_heading_pair (INPUT heading_pair AS CLASS JsonArray):
        
 

        /** Ignore empty pairs such as charactsheets**/
        IF heading_pair:IsNull(1) THEN
            RETURN.
            
        DEFINE VARIABLE tempJsonArray AS CLASS JsonArray .

        tempJsonArray = NEW JsonArray().
        tempJsonArray:ADD( 'lpstr' ).
        tempJsonArray:ADD( heading_pair:GETCHARACTER(1) ).
        THIS-OBJECT:heading_pairs:ADD( tempJsonArray ).
        
        tempJsonArray = NEW JsonArray().
        tempJsonArray:ADD( 'i4' ).
        tempJsonArray:ADD( heading_pair:GETCHARACTER(2) ).
        THIS-OBJECT:heading_pairs:ADD( tempJsonArray ).
        


    END METHOD.
    
    METHOD PUBLIC VOID _set_properties(INPUT property AS CHARACTER,
                                       INPUT propValue AS CHARACTER):
        myAppProp:SET(property, propValue).
    END METHOD.
        
    METHOD PUBLIC VOID _assemble_xml_file():
        

        THIS-OBJECT:_write_properties().

        /** Close the file.**/
        THIS-OBJECT:_xml_close().
    END METHOD.

    METHOD PRIVATE VOID _write_properties():
        
        hnSAXWriter:START-ELEMENT('Properties').

        hnSAXWriter:DECLARE-NAMESPACE('http://schemas.openxmlformats.org/officeDocument/2006/extended-properties').
        hnSAXWriter:DECLARE-NAMESPACE('http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes','vt').

        THIS-OBJECT:_write_application().
        THIS-OBJECT:_write_doc_security().
        THIS-OBJECT:_write_scale_crop().
        THIS-OBJECT:_write_heading_Pairs().
        THIS-OBJECT:_write_title_of_parts().
        THIS-OBJECT:_write_manager().
        THIS-OBJECT:_write_company().
        THIS-OBJECT:_write_links_up_to_date().
        THIS-OBJECT:_write_shared_doc().
        THIS-OBJECT:_write_hyperlinks_changed().
        THIS-OBJECT:_write_app_version().

        hnSAXWriter:END-ELEMENT('Properties').

    END METHOD.

    METHOD PRIVATE VOID _write_application():
        hnSAXWriter:WRITE-DATA-ELEMENT ('Application','Microsoft Excel').
    END METHOD.

    METHOD PRIVATE VOID _write_doc_security():
        hnSAXWriter:WRITE-DATA-ELEMENT ('DocSecurity','0').
    END METHOD.

    METHOD PRIVATE VOID _write_scale_crop():

         hnSAXWriter:WRITE-DATA-ELEMENT ('ScaleCrop','false').
    END METHOD.

    METHOD PRIVATE VOID _write_heading_Pairs():

        hnSAXWriter:START-ELEMENT('HeadingPairs').
        THIS-OBJECT:_write_vt_vector('variant', THIS-OBJECT:heading_pairs).
        
        hnSAXWriter:END-ELEMENT('HeadingPairs').

    END METHOD.

    METHOD PRIVATE VOID _write_title_of_parts():

        hnSAXWriter:START-ELEMENT('TitlesOfParts').
         THIS-OBJECT:_write_vt_vector('lpstr', THIS-OBJECT:part_name).
        hnSAXWriter:END-ELEMENT('TitlesOfParts').
    END METHOD.

    METHOD PRIVATE VOID _write_vt_vector(INPUT base_Type   AS CHARACTER,
                                         INPUT vector_data AS CLASS JsonArray):

        DEFINE VARIABLE inVectorData_Entry AS INTEGER     NO-UNDO.

        hnSAXWriter:START-ELEMENT('vt:vector').    
        hnSAXWriter:INSERT-ATTRIBUTE('size', STRING(vector_data:length)).    
        hnSAXWriter:INSERT-ATTRIBUTE('baseType',  base_Type ).

        DO inVectorData_Entry = 1 TO vector_data:LENGTH:

            IF base_Type EQ 'variant' THEN
                hnSAXWriter:START-ELEMENT('vt:variant').
            
            THIS-OBJECT:_write_vt_data( vector_data:GetJsonArray( inVectorData_Entry ) ).

            IF base_Type EQ 'variant' THEN
                hnSAXWriter:END-ELEMENT('vt:variant').
        END.

        hnSAXWriter:END-ELEMENT('vt:vector').    
    END METHOD.

    METHOD PRIVATE VOID _write_vt_data(INPUT vtData AS JsonArray):
        hnSAXWriter:WRITE-DATA-ELEMENT( 'vt:' + vtData:GETCHARACTER(1), vtData:GETCHARACTER(2) ).
    END METHOD.

    METHOD PRIVATE VOID _write_manager():
    END METHOD.

    METHOD PRIVATE VOID _write_company():
    END METHOD.

    METHOD PRIVATE VOID _write_links_up_to_date():
        hnSAXWriter:WRITE-DATA-ELEMENT ('LinksUpToDate','false').
    END METHOD.

    METHOD PRIVATE VOID _write_shared_doc():
        hnSAXWriter:WRITE-DATA-ELEMENT ('SharedDoc','false').
    END METHOD.

    METHOD PRIVATE VOID _write_hyperlinks_changed():
        hnSAXWriter:WRITE-DATA-ELEMENT ('HyperlinksChanged','false').
    END METHOD.

    METHOD PRIVATE VOID _write_app_version():
        hnSAXWriter:WRITE-DATA-ELEMENT ('AppVersion','14.0300').
    END METHOD.

END CLASS.

