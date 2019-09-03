package html;

import java.math.BigInteger;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;

/**
 * SHA256加密，用来加密文件名，避免特殊字符在不同环境下的乱码
 * */
public class SHA256 {
    public static String getSHA(String input)
    {
        try {
            // SHA静态调用
            MessageDigest md = MessageDigest.getInstance("SHA-256");

            //计算消息摘要，并且返回字节数组
            byte[] messageDigest = md.digest(input.getBytes());

            // 将字节数组转换为signum表示
            BigInteger no = new BigInteger(1, messageDigest);

            // 将消息摘要转换为十六进制值
            String hashtext = no.toString(16);
            // 写入加密结果
            while (hashtext.length() < 32) {
                hashtext = "0" + hashtext;
            }

            return hashtext;
        }
        catch (NoSuchAlgorithmException e) {// 用于指定错误的消息摘要算法
            e.printStackTrace();
            return  null;
        }
    }
}