package html;

import org.apache.commons.codec.binary.Base64;
import org.apache.commons.io.FileUtils;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.xwpf.converter.core.BasicURIResolver;
import org.apache.poi.xwpf.converter.core.FileImageExtractor;
import org.apache.poi.xwpf.converter.xhtml.XHTMLConverter;
import org.apache.poi.xwpf.converter.xhtml.XHTMLOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.w3c.dom.Document;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.*;

public class WordToHtml {
    /**
     * 将word2003转换为html
     *
     * @param wordPath
     *            word文件路径
     * @param wordName
     *            word文件名称无后缀
     * @param suffix
     *            word文件后缀
     * @throws IOException
     * @throws TransformerException
     * @throws ParserConfigurationException
     */
    public String word2003ToHtml(String wordPath, String wordName,
                                 String suffix) throws IOException, TransformerException,
            ParserConfigurationException {
        File directory = new File(".");
        String htmlName = directory.getCanonicalPath() + SHA256.getSHA(wordName) + ".html";

        // 原word文档
        final String file = wordPath + File.separator + wordName + suffix;
        InputStream input = new FileInputStream(new File(file));

        HWPFDocument wordDocument = new HWPFDocument(input);
        WordToHtmlConverter wordToHtmlConverter = new ImageWordToHtmlConverter(
                DocumentBuilderFactory.newInstance().newDocumentBuilder()
                        .newDocument());

        wordToHtmlConverter.processDocument(wordDocument);
        Document htmlDocument = wordToHtmlConverter.getDocument();

        DOMSource domSource = new DOMSource(htmlDocument);
        StreamResult streamResult = new StreamResult(new File(htmlName));
        TransformerFactory tf = TransformerFactory.newInstance();
        Transformer serializer = tf.newTransformer();
        serializer.setOutputProperty(OutputKeys.ENCODING, "utf-8");
        serializer.setOutputProperty(OutputKeys.INDENT, "yes");
        serializer.setOutputProperty(OutputKeys.METHOD, "html");
        serializer.transform(domSource, streamResult);
        File initFile = new File(htmlName);
        FileReader reader = new FileReader(initFile);
        BufferedReader bReader = new BufferedReader(reader);
        StringBuilder sb = new StringBuilder();
        String s = "";
        while ((s =bReader.readLine()) != null) {
            sb.append(s);
        }
        bReader.close();
        FileUtils.deleteQuietly(initFile);
        return sb.toString();
    }

    /**
     * 将word2003转换为html，并且保存html文件
     *
     * @param wordPath
     *            word文件路径
     * @param wordName
     *            word文件名称无后缀
     * @param suffix
     *            word文件后缀
     * @param savePath
     *            html保存路径
     * @throws IOException
     * @throws TransformerException
     * @throws ParserConfigurationException
     */
    public String word2003ToHtml(String wordPath, String wordName,
                                 String suffix,String savePath) throws IOException, TransformerException,
            ParserConfigurationException {
        String htmlName = SHA256.getSHA(wordName) + ".html";
        String htmlPath = savePath + htmlName;

        // 原word文档
        final String file = wordPath + File.separator + wordName + suffix;
        InputStream input = new FileInputStream(new File(file));

        HWPFDocument wordDocument = new HWPFDocument(input);
        WordToHtmlConverter wordToHtmlConverter = new ImageWordToHtmlConverter(
                DocumentBuilderFactory.newInstance().newDocumentBuilder()
                        .newDocument());

        // 判断html文件是否存在，每次重新生成
        File htmlFile = new File(htmlPath);

        // 解析word文档
        wordToHtmlConverter.processDocument(wordDocument);
        Document htmlDocument = wordToHtmlConverter.getDocument();

        // 生成html文件上级文件夹
        File folder = new File(savePath);
        if (!folder.exists()) {
            folder.mkdirs();
        }

        // 生成html文件地址
        OutputStream outStream = new FileOutputStream(htmlFile);

        DOMSource domSource = new DOMSource(htmlDocument);
        StreamResult streamResult = new StreamResult(outStream);

        TransformerFactory factory = TransformerFactory.newInstance();
        Transformer serializer = factory.newTransformer();
        serializer.setOutputProperty(OutputKeys.ENCODING, "utf-8");
        serializer.setOutputProperty(OutputKeys.INDENT, "yes");
        serializer.setOutputProperty(OutputKeys.METHOD, "html");

        serializer.transform(domSource, streamResult);

        outStream.close();

        return htmlName;
    }

    /**
     * 2007版本word转换成html
     *
     * @param wordPath
     *            word文件路径
     * @param wordName
     *            word文件名称无后缀
     * @param suffix
     *            word文件后缀
     * @return
     * @throws IOException
     */
    public String word2007ToHtml(String wordPath, String wordName, String suffix)
            throws IOException {
        XWPFDocument docxDocument = new XWPFDocument(
                new FileInputStream(wordPath + File.separator + wordName + suffix));
        // 配置
        XHTMLOptions options = XHTMLOptions.create();
        // 设置图片存储路径
        String path = System.getProperty("java.io.tmpdir");
        String firstImagePathStr = path + File.separator + String.valueOf(System.currentTimeMillis());
        options.setExtractor(new FileImageExtractor(new File(firstImagePathStr)));
        options.URIResolver(new BasicURIResolver(firstImagePathStr));
        // 转换html
        ByteArrayOutputStream htmlStream = new ByteArrayOutputStream();
        XHTMLConverter.getInstance().convert(docxDocument, htmlStream, options);
        String htmlStr = htmlStream.toString();
        // 将image文件转换为base64并替换到html字符串里
        String middleImageDirStr = "/word/media";
        String imageDirStr = firstImagePathStr + middleImageDirStr;
        File imageDir = new File(imageDirStr);
        String[] imageList = imageDir.list();
        if (imageList != null) {
            for (int i = 0; i < imageList.length; i++) {
                String oneImagePathStr = imageDirStr + "/" + imageList[i];
                File oneImageFile = new File(oneImagePathStr);
                String imageBase64Str = new String(Base64.encodeBase64(FileUtils.readFileToByteArray(oneImageFile)), "UTF-8");
                htmlStr = htmlStr.replace(oneImagePathStr, "data:image/png;base64," + imageBase64Str);
            }
        }
        //删除图片路径
        File firstImagePath = new File(firstImagePathStr);
        FileUtils.deleteDirectory(firstImagePath);
        return htmlStr;
    }

    /**
     * 将word2007转换为html，并且保存html文件
     *
     * @param wordPath
     *            word文件路径
     * @param wordName
     *            word文件名称无后缀
     * @param suffix
     *            word文件后缀
     * @param savePath
     *            html保存路径
     * @throws IOException
     * @throws TransformerException
     * @throws ParserConfigurationException
     */
    public String word2007ToHtml(String wordPath, String wordName, String suffix,String savePath)
            throws IOException {
        String html = word2007ToHtml(wordPath,wordName,suffix);

        byte[] sourceByte = html.getBytes();
        String htmlName = SHA256.getSHA(wordName) + ".html";
        if(null != sourceByte){
            File file = new File(savePath + htmlName);//文件路径（路径+文件名）
            if (!file.exists()) {	//文件不存在则创建文件，先创建目录
                File dir = new File(file.getParent());
                dir.mkdirs();
                file.createNewFile();
            }
            FileOutputStream outStream = new FileOutputStream(file);	//文件输出流用于将数据写入文件
            outStream.write(sourceByte);
            outStream.close();	//关闭文件输出流
        }
        return htmlName;
    }

    public static void main(String[] args) {
        WordToHtml wordToHtml = new WordToHtml();
        try {
            String name = wordToHtml.word2007ToHtml(
                    "F:\\","rtf",".rtf","F:\\测试\\");
            System.out.println(name);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}