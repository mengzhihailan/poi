package pdf;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import com.aspose.cells.Workbook;
import com.aspose.words.Document;
import com.aspose.words.License;

/**
 * Word或Excel 转Pdf 帮助类
 */
public class PdfUtil {
    private static boolean getLicense() {
        boolean result = false;
        try {
            InputStream is = PdfUtil.class.getClassLoader().getResourceAsStream("license.xml"); // license.xml应放在..\WebRoot\WEB-INF\classes路径下
            License aposeLic = new License();
            aposeLic.setLicense(is);
            result = true;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return result;
    }

    /**
     * @param wordPath 需要被转换的word全路径带文件名
     * @param pdfPath 转换之后pdf的全路径带文件名
     */
    public static void doc2pdf(String wordPath, String pdfPath) {
        if (!getLicense()) { // 验证License 若不验证则转化出的pdf文档会有水印产生
            return;
        }
        long old = System.currentTimeMillis();
        File file = new File(pdfPath); //新建一个pdf文档
        try(FileOutputStream os = new FileOutputStream(file)) {
            Document doc = new Document(wordPath); //Address是将要被转化的word文档
            //全面支持DOC, DOCX, OOXML, RTF HTML, OpenDocument, PDF, EPUB, XPS, SWF 相互转换
            doc.save(os, com.aspose.words.SaveFormat.PDF);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * @param excelPath 需要被转换的excel全路径带文件名
     * @param pdfPath 转换之后pdf的全路径带文件名
     */
    public static void excel2pdf(String excelPath, String pdfPath) {
        if (!getLicense()) { // 验证License 若不验证则转化出的pdf文档会有水印产生
            return;
        }
        try {
            long old = System.currentTimeMillis();
            Workbook wb = new Workbook(excelPath);// 原始excel路径
            try( FileOutputStream fileOS = new FileOutputStream(new File(pdfPath))){
                wb.save(fileOS, com.aspose.cells.SaveFormat.PDF);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}