package com.excelparser.parser;

/**
 * See XSSFRichTextString in org.apache.poi.xssf.usermodel. This class mimics the getString,
 * but passes the CTRst object as a parameter so that I only one XSSFRichTextStringParser
 * has to be created and can be reused.
 */
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTRElt;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTRst;

public class XSSFRichTextStringParser {

	private static final Pattern utfPtrn = Pattern.compile("_x([0-9A-F]{4})_");

	@SuppressWarnings("deprecation")
	public String getString(CTRst st) {
        if(st.sizeOfRArray() == 0) {
            return utfDecode(st.getT());
        }
        StringBuffer buf = new StringBuffer();
        for(CTRElt r : st.getRArray()){
            buf.append(r.getT());
        }
        return utfDecode(buf.toString());
    }

    private String utfDecode(String value){
        if(value == null) return null;
        
        StringBuffer buf = new StringBuffer();
        Matcher m = utfPtrn.matcher(value);
        int idx = 0;
        while(m.find()) {
            int pos = m.start();
            if( pos > idx) {
                buf.append(value.substring(idx, pos));
            }

            String code = m.group(1);
            int icode = Integer.decode("0x" + code);
            buf.append((char)icode);

            idx = m.end();
        }
        buf.append(value.substring(idx));
        return buf.toString();
    }

}
