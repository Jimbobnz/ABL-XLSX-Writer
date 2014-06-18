
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
CLASS ABL-XLSX-Writer.ABL-XLSX-Writer.writer.Theme INHERITS xmlwriter:

    METHOD PUBLIC VOID _assemble_xml_file():

        THIS-OBJECT:_write_theme_file().

        /** Close the file.**/
        THIS-OBJECT:_xml_close().

    END METHOD.

    METHOD PUBLIC VOID _write_theme_file():

        IF NOT VALID-HANDLE(hnSAXWriter) OR 
           hnSAXWriter:WRITE-STATUS EQ SAX-WRITE-ERROR THEN
            RETURN.
            
        /** Don't Touch this code it's working to well to mess around with.**/

        
        	hnSAXWriter:START-ELEMENT("a:theme", "http://schemas.openxmlformats.org/drawingml/2006/main").
        		hnSAXWriter:INSERT-ATTRIBUTE("name","Office Theme").
        		hnSAXWriter:START-ELEMENT("a:themeElements", "http://schemas.openxmlformats.org/drawingml/2006/main").
        			hnSAXWriter:START-ELEMENT("a:clrScheme", "http://schemas.openxmlformats.org/drawingml/2006/main").
        				hnSAXWriter:INSERT-ATTRIBUTE("name","Office").
        				hnSAXWriter:START-ELEMENT("a:dk1", "http://schemas.openxmlformats.org/drawingml/2006/main").
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:sysClr", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("val","windowText").
        						hnSAXWriter:INSERT-ATTRIBUTE("lastClr","000000").
        				hnSAXWriter:END-ELEMENT("a:dk1").

        				hnSAXWriter:START-ELEMENT("a:lt1", "http://schemas.openxmlformats.org/drawingml/2006/main").
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:sysClr", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("val","window").
        						hnSAXWriter:INSERT-ATTRIBUTE("lastClr","FFFFFF").
        					
        				hnSAXWriter:END-ELEMENT("a:lt1").

        				hnSAXWriter:START-ELEMENT("a:dk2", "http://schemas.openxmlformats.org/drawingml/2006/main").
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:srgbClr", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("val","1F497D").
        					
        				hnSAXWriter:END-ELEMENT("a:dk2").

        				hnSAXWriter:START-ELEMENT("a:lt2", "http://schemas.openxmlformats.org/drawingml/2006/main").
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:srgbClr", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("val","EEECE1").
        					
        				hnSAXWriter:END-ELEMENT("a:lt2").

        				hnSAXWriter:START-ELEMENT("a:accent1", "http://schemas.openxmlformats.org/drawingml/2006/main").
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:srgbClr", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("val","4F81BD").
        					
        				hnSAXWriter:END-ELEMENT("a:accent1").

        				hnSAXWriter:START-ELEMENT("a:accent2", "http://schemas.openxmlformats.org/drawingml/2006/main").
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:srgbClr", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("val","C0504D").
        					
        				hnSAXWriter:END-ELEMENT("a:accent2").

        				hnSAXWriter:START-ELEMENT("a:accent3", "http://schemas.openxmlformats.org/drawingml/2006/main").
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:srgbClr", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("val","9BBB59").
        					
        				hnSAXWriter:END-ELEMENT("a:accent3").

        				hnSAXWriter:START-ELEMENT("a:accent4", "http://schemas.openxmlformats.org/drawingml/2006/main").
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:srgbClr", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("val","8064A2").
        					
        				hnSAXWriter:END-ELEMENT("a:accent4").

        				hnSAXWriter:START-ELEMENT("a:accent5", "http://schemas.openxmlformats.org/drawingml/2006/main").
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:srgbClr", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("val","4BACC6").
        					
        				hnSAXWriter:END-ELEMENT("a:accent5").

        				hnSAXWriter:START-ELEMENT("a:accent6", "http://schemas.openxmlformats.org/drawingml/2006/main").
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:srgbClr", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("val","F79646").
        					
        				hnSAXWriter:END-ELEMENT("a:accent6").

        				hnSAXWriter:START-ELEMENT("a:hlink", "http://schemas.openxmlformats.org/drawingml/2006/main").
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:srgbClr", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("val","0000FF").
        					
        				hnSAXWriter:END-ELEMENT("a:hlink").

        				hnSAXWriter:START-ELEMENT("a:folHlink", "http://schemas.openxmlformats.org/drawingml/2006/main").
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:srgbClr", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("val","800080").
        					
        				hnSAXWriter:END-ELEMENT("a:folHlink").

        			hnSAXWriter:END-ELEMENT("a:clrScheme").

        			hnSAXWriter:START-ELEMENT("a:fontScheme", "http://schemas.openxmlformats.org/drawingml/2006/main").
        				hnSAXWriter:INSERT-ATTRIBUTE("name","Office").
        				hnSAXWriter:START-ELEMENT("a:majorFont", "http://schemas.openxmlformats.org/drawingml/2006/main").

        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:latin", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface","Cambria").
        					
                            hnSAXWriter:WRITE-EMPTY-ELEMENT("a:ea", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface","").
        					
                            hnSAXWriter:WRITE-EMPTY-ELEMENT("a:cs", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface","").
        					
                            hnSAXWriter:WRITE-EMPTY-ELEMENT("a:font", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("script","Jpan").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface", chr(0xEFBCAD) + chr(0xEFBCB3) + ' ' + chr(0xEFBCB0) + chr(0xE382B4) + chr(0xE382B7) + chr(0xE38383) + chr(0xE382AF)).
        					
                            hnSAXWriter:WRITE-EMPTY-ELEMENT("a:font", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("script","Hang").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface", chr(0xEBA791) + chr(0xEC9D80) + ' ' + chr(0xEAB3A0) + chr(0xEB9495) ).
        					
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:font", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("script","Hans").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface", chr(0xE5AE8B) + chr(0xE4BD93) ).
        					
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:font", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("script","Hant").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface",CHR(0xE696B0) + CHR(0xE7B4B0) + CHR(0xE6988E) + CHR(0xE9AB94) ).
        					
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:font", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("script","Arab").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface","Times New Roman").
        					
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:font", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("script","Hebr").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface","Times New Roman").
        					
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:font", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("script","Thai").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface","Tahoma").
        					
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:font", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("script","Ethi").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface","Nyala").
        					
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:font", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("script","Beng").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface","Vrinda").
        					
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:font", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("script","Gujr").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface","Shruti").
        					
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:font", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("script","Khmr").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface","MoolBoran").
        					
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:font", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("script","Knda").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface","Tunga").
        					
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:font", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("script","Guru").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface","Raavi").
        					
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:font", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("script","Cans").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface","Euphemia").
        					
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:font", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("script","Cher").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface","Plantagenet Cherokee").
        					
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:font", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("script","Yiii").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface","Microsoft Yi Baiti").
        					
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:font", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("script","Tibt").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface","Microsoft Himalaya").
        					
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:font", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("script","Thaa").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface","MV Boli").
        					
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:font", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("script","Deva").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface","Mangal").
        					
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:font", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("script","Telu").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface","Gautami").
        					
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:font", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("script","Taml").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface","Latha").
        					
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:font", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("script","Syrc").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface","Estrangelo Edessa").
        					
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:font", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("script","Orya").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface","Kalinga").
        					
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:font", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("script","Mlym").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface","Kartika").
        					
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:font", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("script","Laoo").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface","DokChampa").
        					
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:font", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("script","Sinh").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface","Iskoola Pota").
        					
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:font", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("script","Mong").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface","Mongolian Baiti").
        					
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:font", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("script","Viet").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface","Times New Roman").
        					
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:font", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("script","Uigh").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface","Microsoft Uighur").
        					
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:font", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("script","Geor").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface","Sylfaen").
        					
        				hnSAXWriter:END-ELEMENT("a:majorFont").
        				hnSAXWriter:START-ELEMENT("a:minorFont", "http://schemas.openxmlformats.org/drawingml/2006/main").
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:latin", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface","Calibri").
        					
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:ea", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface","").
        					
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:cs", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface","").
        					
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:font", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("script","Jpan").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface", chr(0xEFBCAD) + chr(0xEFBCB3) + ' ' + chr(0xEFBCB0) + chr(0xE382B4) + chr(0xE382B7) + chr(0xE38383) + chr(0xE382AF)).
        					
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:font", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("script","Hang").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface", chr(0xEBA791) + chr(0xEC9D80) + ' ' + chr(0xEAB3A0) + chr(0xEB9495) ).
        					
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:font", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("script","Hans").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface", chr(0xE5AE8B) + chr(0xE4BD93) ).
        					
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:font", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("script","Hant").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface",CHR(0xE696B0) + CHR(0xE7B4B0) + CHR(0xE6988E) + CHR(0xE9AB94) ).
        					
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:font", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("script","Arab").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface","Arial").
        					
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:font", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("script","Hebr").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface","Arial").
        					
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:font", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("script","Thai").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface","Tahoma").
        					
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:font", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("script","Ethi").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface","Nyala").
        					
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:font", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("script","Beng").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface","Vrinda").
        					
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:font", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("script","Gujr").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface","Shruti").
        					
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:font", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("script","Khmr").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface","DaunPenh").
        					
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:font", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("script","Knda").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface","Tunga").
        					
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:font", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("script","Guru").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface","Raavi").
        					
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:font", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("script","Cans").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface","Euphemia").
        					
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:font", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("script","Cher").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface","Plantagenet Cherokee").
        					
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:font", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("script","Yiii").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface","Microsoft Yi Baiti").
        					
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:font", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("script","Tibt").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface","Microsoft Himalaya").
        					
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:font", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("script","Thaa").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface","MV Boli").
        					
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:font", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("script","Deva").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface","Mangal").
        					
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:font", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("script","Telu").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface","Gautami").
        					
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:font", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("script","Taml").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface","Latha").
        					
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:font", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("script","Syrc").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface","Estrangelo Edessa").
        					
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:font", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("script","Orya").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface","Kalinga").
        					
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:font", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("script","Mlym").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface","Kartika").
        					
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:font", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("script","Laoo").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface","DokChampa").
        					
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:font", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("script","Sinh").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface","Iskoola Pota").
        					
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:font", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("script","Mong").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface","Mongolian Baiti").
        					
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:font", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("script","Viet").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface","Arial").
        					
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:font", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("script","Uigh").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface","Microsoft Uighur").
        					
        					hnSAXWriter:WRITE-EMPTY-ELEMENT("a:font", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("script","Geor").
        						hnSAXWriter:INSERT-ATTRIBUTE("typeface","Sylfaen").
        					
        				hnSAXWriter:END-ELEMENT("a:minorFont").
        			hnSAXWriter:END-ELEMENT("a:fontScheme").
        			hnSAXWriter:START-ELEMENT("a:fmtScheme", "http://schemas.openxmlformats.org/drawingml/2006/main").
        				hnSAXWriter:INSERT-ATTRIBUTE("name","Office").
        				hnSAXWriter:START-ELEMENT("a:fillStyleLst", "http://schemas.openxmlformats.org/drawingml/2006/main").
        					hnSAXWriter:START-ELEMENT("a:solidFill", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:WRITE-EMPTY-ELEMENT("a:schemeClr", "http://schemas.openxmlformats.org/drawingml/2006/main").
        							hnSAXWriter:INSERT-ATTRIBUTE("val","phClr").
        						
        					hnSAXWriter:END-ELEMENT("a:solidFill").
        					hnSAXWriter:START-ELEMENT("a:gradFill", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("rotWithShape","1").
        						hnSAXWriter:START-ELEMENT("a:gsLst", "http://schemas.openxmlformats.org/drawingml/2006/main").
        							hnSAXWriter:START-ELEMENT("a:gs", "http://schemas.openxmlformats.org/drawingml/2006/main").
        								hnSAXWriter:INSERT-ATTRIBUTE("pos","0").
        								hnSAXWriter:START-ELEMENT("a:schemeClr", "http://schemas.openxmlformats.org/drawingml/2006/main").
        									hnSAXWriter:INSERT-ATTRIBUTE("val","phClr").
        									hnSAXWriter:WRITE-EMPTY-ELEMENT("a:tint", "http://schemas.openxmlformats.org/drawingml/2006/main").
        										hnSAXWriter:INSERT-ATTRIBUTE("val","50000").
        									
        									hnSAXWriter:WRITE-EMPTY-ELEMENT("a:satMod", "http://schemas.openxmlformats.org/drawingml/2006/main").
        										hnSAXWriter:INSERT-ATTRIBUTE("val","300000").
        									
        								hnSAXWriter:END-ELEMENT("a:schemeClr").
        							hnSAXWriter:END-ELEMENT("a:gs").
        							hnSAXWriter:START-ELEMENT("a:gs", "http://schemas.openxmlformats.org/drawingml/2006/main").
        								hnSAXWriter:INSERT-ATTRIBUTE("pos","35000").
        								hnSAXWriter:START-ELEMENT("a:schemeClr", "http://schemas.openxmlformats.org/drawingml/2006/main").
        									hnSAXWriter:INSERT-ATTRIBUTE("val","phClr").
        									hnSAXWriter:WRITE-EMPTY-ELEMENT("a:tint", "http://schemas.openxmlformats.org/drawingml/2006/main").
        										hnSAXWriter:INSERT-ATTRIBUTE("val","37000").
        									
        									hnSAXWriter:WRITE-EMPTY-ELEMENT("a:satMod", "http://schemas.openxmlformats.org/drawingml/2006/main").
        										hnSAXWriter:INSERT-ATTRIBUTE("val","300000").
        									
        								hnSAXWriter:END-ELEMENT("a:schemeClr").
        							hnSAXWriter:END-ELEMENT("a:gs").
        							hnSAXWriter:START-ELEMENT("a:gs", "http://schemas.openxmlformats.org/drawingml/2006/main").
        								hnSAXWriter:INSERT-ATTRIBUTE("pos","100000").
        								hnSAXWriter:START-ELEMENT("a:schemeClr", "http://schemas.openxmlformats.org/drawingml/2006/main").
        									hnSAXWriter:INSERT-ATTRIBUTE("val","phClr").
        									hnSAXWriter:WRITE-EMPTY-ELEMENT("a:tint", "http://schemas.openxmlformats.org/drawingml/2006/main").
        										hnSAXWriter:INSERT-ATTRIBUTE("val","15000").
        									
        									hnSAXWriter:WRITE-EMPTY-ELEMENT("a:satMod", "http://schemas.openxmlformats.org/drawingml/2006/main").
        										hnSAXWriter:INSERT-ATTRIBUTE("val","350000").
        									
        								hnSAXWriter:END-ELEMENT("a:schemeClr").
        							hnSAXWriter:END-ELEMENT("a:gs").
        						hnSAXWriter:END-ELEMENT("a:gsLst").
        						hnSAXWriter:WRITE-EMPTY-ELEMENT("a:lin", "http://schemas.openxmlformats.org/drawingml/2006/main").
        							hnSAXWriter:INSERT-ATTRIBUTE("ang","16200000").
        							hnSAXWriter:INSERT-ATTRIBUTE("scaled","1").
        						
        					hnSAXWriter:END-ELEMENT("a:gradFill").
        					hnSAXWriter:START-ELEMENT("a:gradFill", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("rotWithShape","1").
        						hnSAXWriter:START-ELEMENT("a:gsLst", "http://schemas.openxmlformats.org/drawingml/2006/main").
        							hnSAXWriter:START-ELEMENT("a:gs", "http://schemas.openxmlformats.org/drawingml/2006/main").
        								hnSAXWriter:INSERT-ATTRIBUTE("pos","0").
        								hnSAXWriter:START-ELEMENT("a:schemeClr", "http://schemas.openxmlformats.org/drawingml/2006/main").
        									hnSAXWriter:INSERT-ATTRIBUTE("val","phClr").
        									hnSAXWriter:WRITE-EMPTY-ELEMENT("a:shade", "http://schemas.openxmlformats.org/drawingml/2006/main").
        										hnSAXWriter:INSERT-ATTRIBUTE("val","51000").
        									
        									hnSAXWriter:WRITE-EMPTY-ELEMENT("a:satMod", "http://schemas.openxmlformats.org/drawingml/2006/main").
        										hnSAXWriter:INSERT-ATTRIBUTE("val","130000").
        									
        								hnSAXWriter:END-ELEMENT("a:schemeClr").
        							hnSAXWriter:END-ELEMENT("a:gs").
        							hnSAXWriter:START-ELEMENT("a:gs", "http://schemas.openxmlformats.org/drawingml/2006/main").
        								hnSAXWriter:INSERT-ATTRIBUTE("pos","80000").
        								hnSAXWriter:START-ELEMENT("a:schemeClr", "http://schemas.openxmlformats.org/drawingml/2006/main").
        									hnSAXWriter:INSERT-ATTRIBUTE("val","phClr").
        									hnSAXWriter:WRITE-EMPTY-ELEMENT("a:shade", "http://schemas.openxmlformats.org/drawingml/2006/main").
        										hnSAXWriter:INSERT-ATTRIBUTE("val","93000").
        									
        									hnSAXWriter:WRITE-EMPTY-ELEMENT("a:satMod", "http://schemas.openxmlformats.org/drawingml/2006/main").
        										hnSAXWriter:INSERT-ATTRIBUTE("val","130000").
        									
        								hnSAXWriter:END-ELEMENT("a:schemeClr").
        							hnSAXWriter:END-ELEMENT("a:gs").
        							hnSAXWriter:START-ELEMENT("a:gs", "http://schemas.openxmlformats.org/drawingml/2006/main").
        								hnSAXWriter:INSERT-ATTRIBUTE("pos","100000").
        								hnSAXWriter:START-ELEMENT("a:schemeClr", "http://schemas.openxmlformats.org/drawingml/2006/main").
        									hnSAXWriter:INSERT-ATTRIBUTE("val","phClr").
        									hnSAXWriter:WRITE-EMPTY-ELEMENT("a:shade", "http://schemas.openxmlformats.org/drawingml/2006/main").
        										hnSAXWriter:INSERT-ATTRIBUTE("val","94000").
        									
        									hnSAXWriter:WRITE-EMPTY-ELEMENT("a:satMod", "http://schemas.openxmlformats.org/drawingml/2006/main").
        										hnSAXWriter:INSERT-ATTRIBUTE("val","135000").
        									
        								hnSAXWriter:END-ELEMENT("a:schemeClr").
        							hnSAXWriter:END-ELEMENT("a:gs").
        						hnSAXWriter:END-ELEMENT("a:gsLst").
        						hnSAXWriter:WRITE-EMPTY-ELEMENT("a:lin", "http://schemas.openxmlformats.org/drawingml/2006/main").
        							hnSAXWriter:INSERT-ATTRIBUTE("ang","16200000").
        							hnSAXWriter:INSERT-ATTRIBUTE("scaled","0").
        						
        					hnSAXWriter:END-ELEMENT("a:gradFill").
        				hnSAXWriter:END-ELEMENT("a:fillStyleLst").
        				hnSAXWriter:START-ELEMENT("a:lnStyleLst", "http://schemas.openxmlformats.org/drawingml/2006/main").
        					hnSAXWriter:START-ELEMENT("a:ln", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("w","9525").
        						hnSAXWriter:INSERT-ATTRIBUTE("cap","flat").
        						hnSAXWriter:INSERT-ATTRIBUTE("cmpd","sng").
        						hnSAXWriter:INSERT-ATTRIBUTE("algn","ctr").
        						hnSAXWriter:START-ELEMENT("a:solidFill", "http://schemas.openxmlformats.org/drawingml/2006/main").
        							hnSAXWriter:START-ELEMENT("a:schemeClr", "http://schemas.openxmlformats.org/drawingml/2006/main").
        								hnSAXWriter:INSERT-ATTRIBUTE("val","phClr").
        								hnSAXWriter:WRITE-EMPTY-ELEMENT("a:shade", "http://schemas.openxmlformats.org/drawingml/2006/main").
        									hnSAXWriter:INSERT-ATTRIBUTE("val","95000").
        								
        								hnSAXWriter:WRITE-EMPTY-ELEMENT("a:satMod", "http://schemas.openxmlformats.org/drawingml/2006/main").
        									hnSAXWriter:INSERT-ATTRIBUTE("val","105000").
        								
        							hnSAXWriter:END-ELEMENT("a:schemeClr").
        						hnSAXWriter:END-ELEMENT("a:solidFill").
        						hnSAXWriter:WRITE-EMPTY-ELEMENT("a:prstDash", "http://schemas.openxmlformats.org/drawingml/2006/main").
        							hnSAXWriter:INSERT-ATTRIBUTE("val","solid").
        						
        					hnSAXWriter:END-ELEMENT("a:ln").
        					hnSAXWriter:START-ELEMENT("a:ln", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("w","25400").
        						hnSAXWriter:INSERT-ATTRIBUTE("cap","flat").
        						hnSAXWriter:INSERT-ATTRIBUTE("cmpd","sng").
        						hnSAXWriter:INSERT-ATTRIBUTE("algn","ctr").
        						hnSAXWriter:START-ELEMENT("a:solidFill", "http://schemas.openxmlformats.org/drawingml/2006/main").
        							hnSAXWriter:WRITE-EMPTY-ELEMENT("a:schemeClr", "http://schemas.openxmlformats.org/drawingml/2006/main").
        								hnSAXWriter:INSERT-ATTRIBUTE("val","phClr").
        							
        						hnSAXWriter:END-ELEMENT("a:solidFill").
        						hnSAXWriter:WRITE-EMPTY-ELEMENT("a:prstDash", "http://schemas.openxmlformats.org/drawingml/2006/main").
        							hnSAXWriter:INSERT-ATTRIBUTE("val","solid").
        						
        					hnSAXWriter:END-ELEMENT("a:ln").
        					hnSAXWriter:START-ELEMENT("a:ln", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("w","38100").
        						hnSAXWriter:INSERT-ATTRIBUTE("cap","flat").
        						hnSAXWriter:INSERT-ATTRIBUTE("cmpd","sng").
        						hnSAXWriter:INSERT-ATTRIBUTE("algn","ctr").
        						hnSAXWriter:START-ELEMENT("a:solidFill", "http://schemas.openxmlformats.org/drawingml/2006/main").
        							hnSAXWriter:WRITE-EMPTY-ELEMENT("a:schemeClr", "http://schemas.openxmlformats.org/drawingml/2006/main").
        								hnSAXWriter:INSERT-ATTRIBUTE("val","phClr").
        							
        						hnSAXWriter:END-ELEMENT("a:solidFill").
        						hnSAXWriter:WRITE-EMPTY-ELEMENT("a:prstDash", "http://schemas.openxmlformats.org/drawingml/2006/main").
        							hnSAXWriter:INSERT-ATTRIBUTE("val","solid").
        						
        					hnSAXWriter:END-ELEMENT("a:ln").
        				hnSAXWriter:END-ELEMENT("a:lnStyleLst").
        				hnSAXWriter:START-ELEMENT("a:effectStyleLst", "http://schemas.openxmlformats.org/drawingml/2006/main").
        					hnSAXWriter:START-ELEMENT("a:effectStyle", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:START-ELEMENT("a:effectLst", "http://schemas.openxmlformats.org/drawingml/2006/main").
        							hnSAXWriter:START-ELEMENT("a:outerShdw", "http://schemas.openxmlformats.org/drawingml/2006/main").
        								hnSAXWriter:INSERT-ATTRIBUTE("blurRad","40000").
        								hnSAXWriter:INSERT-ATTRIBUTE("dist","20000").
        								hnSAXWriter:INSERT-ATTRIBUTE("dir","5400000").
        								hnSAXWriter:INSERT-ATTRIBUTE("rotWithShape","0").
        								hnSAXWriter:START-ELEMENT("a:srgbClr", "http://schemas.openxmlformats.org/drawingml/2006/main").
        									hnSAXWriter:INSERT-ATTRIBUTE("val","000000").
        									hnSAXWriter:WRITE-EMPTY-ELEMENT("a:alpha", "http://schemas.openxmlformats.org/drawingml/2006/main").
        										hnSAXWriter:INSERT-ATTRIBUTE("val","38000").
        									
        								hnSAXWriter:END-ELEMENT("a:srgbClr").
        							hnSAXWriter:END-ELEMENT("a:outerShdw").
        						hnSAXWriter:END-ELEMENT("a:effectLst").
        					hnSAXWriter:END-ELEMENT("a:effectStyle").
        					hnSAXWriter:START-ELEMENT("a:effectStyle", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:START-ELEMENT("a:effectLst", "http://schemas.openxmlformats.org/drawingml/2006/main").
        							hnSAXWriter:START-ELEMENT("a:outerShdw", "http://schemas.openxmlformats.org/drawingml/2006/main").
        								hnSAXWriter:INSERT-ATTRIBUTE("blurRad","40000").
        								hnSAXWriter:INSERT-ATTRIBUTE("dist","23000").
        								hnSAXWriter:INSERT-ATTRIBUTE("dir","5400000").
        								hnSAXWriter:INSERT-ATTRIBUTE("rotWithShape","0").
        								hnSAXWriter:START-ELEMENT("a:srgbClr", "http://schemas.openxmlformats.org/drawingml/2006/main").
        									hnSAXWriter:INSERT-ATTRIBUTE("val","000000").
        									hnSAXWriter:WRITE-EMPTY-ELEMENT("a:alpha", "http://schemas.openxmlformats.org/drawingml/2006/main").
        										hnSAXWriter:INSERT-ATTRIBUTE("val","35000").
        									
        								hnSAXWriter:END-ELEMENT("a:srgbClr").
        							hnSAXWriter:END-ELEMENT("a:outerShdw").
        						hnSAXWriter:END-ELEMENT("a:effectLst").
        					hnSAXWriter:END-ELEMENT("a:effectStyle").
        					hnSAXWriter:START-ELEMENT("a:effectStyle", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:START-ELEMENT("a:effectLst", "http://schemas.openxmlformats.org/drawingml/2006/main").
        							hnSAXWriter:START-ELEMENT("a:outerShdw", "http://schemas.openxmlformats.org/drawingml/2006/main").
        								hnSAXWriter:INSERT-ATTRIBUTE("blurRad","40000").
        								hnSAXWriter:INSERT-ATTRIBUTE("dist","23000").
        								hnSAXWriter:INSERT-ATTRIBUTE("dir","5400000").
        								hnSAXWriter:INSERT-ATTRIBUTE("rotWithShape","0").
        								hnSAXWriter:START-ELEMENT("a:srgbClr", "http://schemas.openxmlformats.org/drawingml/2006/main").
        									hnSAXWriter:INSERT-ATTRIBUTE("val","000000").
        									hnSAXWriter:WRITE-EMPTY-ELEMENT("a:alpha", "http://schemas.openxmlformats.org/drawingml/2006/main").
        										hnSAXWriter:INSERT-ATTRIBUTE("val","35000").
        									
        								hnSAXWriter:END-ELEMENT("a:srgbClr").
        							hnSAXWriter:END-ELEMENT("a:outerShdw").
        						hnSAXWriter:END-ELEMENT("a:effectLst").
        						hnSAXWriter:START-ELEMENT("a:scene3d", "http://schemas.openxmlformats.org/drawingml/2006/main").
        							hnSAXWriter:START-ELEMENT("a:camera", "http://schemas.openxmlformats.org/drawingml/2006/main").
        								hnSAXWriter:INSERT-ATTRIBUTE("prst","orthographicFront").
        								hnSAXWriter:WRITE-EMPTY-ELEMENT("a:rot", "http://schemas.openxmlformats.org/drawingml/2006/main").
        									hnSAXWriter:INSERT-ATTRIBUTE("lat","0").
        									hnSAXWriter:INSERT-ATTRIBUTE("lon","0").
        									hnSAXWriter:INSERT-ATTRIBUTE("rev","0").
        							hnSAXWriter:END-ELEMENT("a:camera").

        							hnSAXWriter:START-ELEMENT("a:lightRig", "http://schemas.openxmlformats.org/drawingml/2006/main").
        								hnSAXWriter:INSERT-ATTRIBUTE("rig","threePt").
        								hnSAXWriter:INSERT-ATTRIBUTE("dir","t").
        								hnSAXWriter:WRITE-EMPTY-ELEMENT("a:rot", "http://schemas.openxmlformats.org/drawingml/2006/main").
        									hnSAXWriter:INSERT-ATTRIBUTE("lat","0").
        									hnSAXWriter:INSERT-ATTRIBUTE("lon","0").
        									hnSAXWriter:INSERT-ATTRIBUTE("rev","1200000").
        								
        							hnSAXWriter:END-ELEMENT("a:lightRig").
        						hnSAXWriter:END-ELEMENT("a:scene3d").
        						hnSAXWriter:START-ELEMENT("a:sp3d", "http://schemas.openxmlformats.org/drawingml/2006/main").
        							hnSAXWriter:WRITE-EMPTY-ELEMENT("a:bevelT", "http://schemas.openxmlformats.org/drawingml/2006/main").
        								hnSAXWriter:INSERT-ATTRIBUTE("w","63500").
        								hnSAXWriter:INSERT-ATTRIBUTE("h","25400").
        							
        						hnSAXWriter:END-ELEMENT("a:sp3d").
        					hnSAXWriter:END-ELEMENT("a:effectStyle").
        				hnSAXWriter:END-ELEMENT("a:effectStyleLst").
        				hnSAXWriter:START-ELEMENT("a:bgFillStyleLst", "http://schemas.openxmlformats.org/drawingml/2006/main").
        					hnSAXWriter:START-ELEMENT("a:solidFill", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:WRITE-EMPTY-ELEMENT("a:schemeClr", "http://schemas.openxmlformats.org/drawingml/2006/main").
        							hnSAXWriter:INSERT-ATTRIBUTE("val","phClr").
        					hnSAXWriter:END-ELEMENT("a:solidFill").

        					hnSAXWriter:START-ELEMENT("a:gradFill", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("rotWithShape","1").
        						hnSAXWriter:START-ELEMENT("a:gsLst", "http://schemas.openxmlformats.org/drawingml/2006/main").
        							hnSAXWriter:START-ELEMENT("a:gs", "http://schemas.openxmlformats.org/drawingml/2006/main").
        								hnSAXWriter:INSERT-ATTRIBUTE("pos","0").
        								hnSAXWriter:START-ELEMENT("a:schemeClr", "http://schemas.openxmlformats.org/drawingml/2006/main").
        									hnSAXWriter:INSERT-ATTRIBUTE("val","phClr").
        									hnSAXWriter:WRITE-EMPTY-ELEMENT("a:tint", "http://schemas.openxmlformats.org/drawingml/2006/main").
        										hnSAXWriter:INSERT-ATTRIBUTE("val","40000").
        									
        									hnSAXWriter:WRITE-EMPTY-ELEMENT("a:satMod", "http://schemas.openxmlformats.org/drawingml/2006/main").
        										hnSAXWriter:INSERT-ATTRIBUTE("val","350000").
        									
        								hnSAXWriter:END-ELEMENT("a:schemeClr").
        							hnSAXWriter:END-ELEMENT("a:gs").
        							hnSAXWriter:START-ELEMENT("a:gs", "http://schemas.openxmlformats.org/drawingml/2006/main").
        								hnSAXWriter:INSERT-ATTRIBUTE("pos","40000").
        								hnSAXWriter:START-ELEMENT("a:schemeClr", "http://schemas.openxmlformats.org/drawingml/2006/main").
        									hnSAXWriter:INSERT-ATTRIBUTE("val","phClr").
        									hnSAXWriter:WRITE-EMPTY-ELEMENT("a:tint", "http://schemas.openxmlformats.org/drawingml/2006/main").
        										hnSAXWriter:INSERT-ATTRIBUTE("val","45000").
        									
        									hnSAXWriter:WRITE-EMPTY-ELEMENT("a:shade", "http://schemas.openxmlformats.org/drawingml/2006/main").
        										hnSAXWriter:INSERT-ATTRIBUTE("val","99000").
        									
        									hnSAXWriter:WRITE-EMPTY-ELEMENT("a:satMod", "http://schemas.openxmlformats.org/drawingml/2006/main").
        										hnSAXWriter:INSERT-ATTRIBUTE("val","350000").
        									
        								hnSAXWriter:END-ELEMENT("a:schemeClr").
        							hnSAXWriter:END-ELEMENT("a:gs").
        							hnSAXWriter:START-ELEMENT("a:gs", "http://schemas.openxmlformats.org/drawingml/2006/main").
        								hnSAXWriter:INSERT-ATTRIBUTE("pos","100000").
        								hnSAXWriter:START-ELEMENT("a:schemeClr", "http://schemas.openxmlformats.org/drawingml/2006/main").
        									hnSAXWriter:INSERT-ATTRIBUTE("val","phClr").
        									hnSAXWriter:WRITE-EMPTY-ELEMENT("a:shade", "http://schemas.openxmlformats.org/drawingml/2006/main").
        										hnSAXWriter:INSERT-ATTRIBUTE("val","20000").
        									
        									hnSAXWriter:WRITE-EMPTY-ELEMENT("a:satMod", "http://schemas.openxmlformats.org/drawingml/2006/main").
        										hnSAXWriter:INSERT-ATTRIBUTE("val","255000").
        									
        								hnSAXWriter:END-ELEMENT("a:schemeClr").
        							hnSAXWriter:END-ELEMENT("a:gs").
        						hnSAXWriter:END-ELEMENT("a:gsLst").
        						hnSAXWriter:START-ELEMENT("a:path", "http://schemas.openxmlformats.org/drawingml/2006/main").
        							hnSAXWriter:INSERT-ATTRIBUTE("path","circle").
        							hnSAXWriter:WRITE-EMPTY-ELEMENT("a:fillToRect", "http://schemas.openxmlformats.org/drawingml/2006/main").
        								hnSAXWriter:INSERT-ATTRIBUTE("l","50000").
        								hnSAXWriter:INSERT-ATTRIBUTE("t","-80000").
        								hnSAXWriter:INSERT-ATTRIBUTE("r","50000").
        								hnSAXWriter:INSERT-ATTRIBUTE("b","180000").
        							
        						hnSAXWriter:END-ELEMENT("a:path").
        					hnSAXWriter:END-ELEMENT("a:gradFill").
        					hnSAXWriter:START-ELEMENT("a:gradFill", "http://schemas.openxmlformats.org/drawingml/2006/main").
        						hnSAXWriter:INSERT-ATTRIBUTE("rotWithShape","1").
        						hnSAXWriter:START-ELEMENT("a:gsLst", "http://schemas.openxmlformats.org/drawingml/2006/main").
        							hnSAXWriter:START-ELEMENT("a:gs", "http://schemas.openxmlformats.org/drawingml/2006/main").
        								hnSAXWriter:INSERT-ATTRIBUTE("pos","0").
        								hnSAXWriter:START-ELEMENT("a:schemeClr", "http://schemas.openxmlformats.org/drawingml/2006/main").
        									hnSAXWriter:INSERT-ATTRIBUTE("val","phClr").
        									hnSAXWriter:WRITE-EMPTY-ELEMENT("a:tint", "http://schemas.openxmlformats.org/drawingml/2006/main").
        										hnSAXWriter:INSERT-ATTRIBUTE("val","80000").
        									
        									hnSAXWriter:WRITE-EMPTY-ELEMENT("a:satMod", "http://schemas.openxmlformats.org/drawingml/2006/main").
        										hnSAXWriter:INSERT-ATTRIBUTE("val","300000").
        									
        								hnSAXWriter:END-ELEMENT("a:schemeClr").
        							hnSAXWriter:END-ELEMENT("a:gs").
        							hnSAXWriter:START-ELEMENT("a:gs", "http://schemas.openxmlformats.org/drawingml/2006/main").
        								hnSAXWriter:INSERT-ATTRIBUTE("pos","100000").
        								hnSAXWriter:START-ELEMENT("a:schemeClr", "http://schemas.openxmlformats.org/drawingml/2006/main").
        									hnSAXWriter:INSERT-ATTRIBUTE("val","phClr").
        									hnSAXWriter:WRITE-EMPTY-ELEMENT("a:shade", "http://schemas.openxmlformats.org/drawingml/2006/main").
        										hnSAXWriter:INSERT-ATTRIBUTE("val","30000").
        									
        									hnSAXWriter:WRITE-EMPTY-ELEMENT("a:satMod", "http://schemas.openxmlformats.org/drawingml/2006/main").
        										hnSAXWriter:INSERT-ATTRIBUTE("val","200000").
        									
        								hnSAXWriter:END-ELEMENT("a:schemeClr").
        							hnSAXWriter:END-ELEMENT("a:gs").
        						hnSAXWriter:END-ELEMENT("a:gsLst").
        						hnSAXWriter:START-ELEMENT("a:path", "http://schemas.openxmlformats.org/drawingml/2006/main").
        							hnSAXWriter:INSERT-ATTRIBUTE("path","circle").
        							hnSAXWriter:WRITE-EMPTY-ELEMENT("a:fillToRect", "http://schemas.openxmlformats.org/drawingml/2006/main").
        								hnSAXWriter:INSERT-ATTRIBUTE("l","50000").
        								hnSAXWriter:INSERT-ATTRIBUTE("t","50000").
        								hnSAXWriter:INSERT-ATTRIBUTE("r","50000").
        								hnSAXWriter:INSERT-ATTRIBUTE("b","50000").
        							
        						hnSAXWriter:END-ELEMENT("a:path").
        					hnSAXWriter:END-ELEMENT("a:gradFill").
        				hnSAXWriter:END-ELEMENT("a:bgFillStyleLst").
        			hnSAXWriter:END-ELEMENT("a:fmtScheme").
        		hnSAXWriter:END-ELEMENT("a:themeElements").
        		hnSAXWriter:WRITE-EMPTY-ELEMENT("a:objectDefaults", "http://schemas.openxmlformats.org/drawingml/2006/main").
        		
        		hnSAXWriter:WRITE-EMPTY-ELEMENT("a:extraClrSchemeLst", "http://schemas.openxmlformats.org/drawingml/2006/main").
        		
        	hnSAXWriter:END-ELEMENT("a:theme").
        
    END METHOD.

END CLASS.
