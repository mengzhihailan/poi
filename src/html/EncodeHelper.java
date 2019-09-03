package html;

import java.io.BufferedInputStream;
import java.io.IOException;

public class EncodeHelper {
    /**
     * 判断文件的编码格式
     * @param bin
     * @return 文件编码格式
     * @throws Exception
     */
    public static String codeString(BufferedInputStream bin) throws IOException {
        int p = (bin.read() << 8) + bin.read();
        String code;
        //其中的 0xefbb、0xfffe、0xfeff、0x5c75这些都是这个文件的前面两个字节的16进制数
        switch (p) {
            case 0xefbb:
                code = "UTF-8";
                break;
            case 0xfffe:
                code = "Unicode";
                break;
            case 0xfeff:
                code = "UTF-16BE";
                break;
            case 0x5c75:
                code = "ANSI|ASCII" ;
                break ;
            default:
                code = "GBK";
        }

        return code;
    }
}
