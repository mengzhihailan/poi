package html;

import java.io.*;

public class TxtToHtml {
    public String toHtml(String sourcePath,String sourceFileName,String suffix,String savePath){
        File file = new File(sourcePath + File.separator + sourceFileName + suffix);
        if(!file.exists())
            file.mkdirs();
        String htmlName = SHA256.getSHA(sourceFileName) + ".html";
        try(BufferedInputStream bis = new BufferedInputStream(new FileInputStream(file));
            InputStreamReader fsr = new InputStreamReader(bis, EncodeHelper.codeString(bis));
            BufferedReader br = new BufferedReader(fsr)){
            StringBuffer content = new StringBuffer();
            String temp;
            temp = br.readLine();
            while (temp != null) {
                temp = br.readLine();
                if(temp != null && !"".equals(temp.trim()))
                    content.append(temp).append("</br>");
            }
            String txtContent = content.toString();
            if(null != txtContent && !"".equals(txtContent)){
                try(FileWriter fw = new FileWriter(savePath + File.separator + htmlName);
                    BufferedWriter bw = new BufferedWriter(fw)){
                    bw.write(txtContent.toCharArray());
                    return htmlName;
                }
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }
}
