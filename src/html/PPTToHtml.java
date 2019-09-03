package html;

import org.apache.poi.hslf.usermodel.HSLFSlide;
import org.apache.poi.hslf.usermodel.HSLFSlideShow;
import org.apache.poi.hslf.usermodel.HSLFTextParagraph;
import org.apache.poi.hslf.usermodel.HSLFTextRun;
import org.apache.poi.xslf.usermodel.*;

import javax.imageio.ImageIO;
import javax.xml.bind.DatatypeConverter;
import java.awt.*;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.*;
import java.util.ArrayList;
import java.util.List;
import java.util.UUID;

public class PPTToHtml {
    /**
     * 将PPT 文件转换成image
     *
     * @param sourcePath PPT文件路径
     * @param htmlSavePath 生成的html存储的文件夹路径
     * @param multiple 图片放大的倍数，提高清晰度
     * @return String html文件的名称
     */
    public String PPT2003toHtml(String sourcePath, String htmlSavePath,int multiple) {
        List<String> base64List = new ArrayList<>();
        String htmlName = SHA256.getSHA(sourcePath) + ".html";
        File fileDir = new File(htmlSavePath + File.separator + htmlName);
        // 如果已经存在，直接返回
        if (fileDir.exists()) {
           return htmlName;
        }
        try(FileInputStream fis = new FileInputStream(sourcePath)) {
            try(HSLFSlideShow hss = new HSLFSlideShow(fis)) {
                // 获取PPT每页的大小（宽和高度）
                Dimension onePPTPageSize = hss.getPageSize();
                // 获得PPT文件中的所有的PPT页面（获得每一张幻灯片）,并转为一张张的播放片
                List<HSLFSlide> slideList = hss.getSlides();
                // 对PPT文件中的每一张幻灯片进行转换和操作
                for (HSLFSlide slide:slideList) {
                    // 设置字体为宋体，防止中文乱码
                    List<List<HSLFTextParagraph>> oneTextParagraphs = slide.getTextParagraphs();
                    for (List<HSLFTextParagraph> list : oneTextParagraphs) {
                        for (HSLFTextParagraph hslfTextParagraph : list) {
                            List<HSLFTextRun> HSLFTextRunList = hslfTextParagraph.getTextRuns();
                            for (int j = 0; j < HSLFTextRunList.size(); j++) {
                                // 如果PPT在WPS中保存过，则
                                // HSLFTextRunList.get(j).getFontSize();的值为0或者26040，
                                // 因此首先识别当前文本框内的字体尺寸是否为0或者大于26040，则设置默认的字体尺寸。

                                // 设置字体大小
                                Double size = HSLFTextRunList.get(j).getFontSize();
                                if ((size <= 0) || (size >= 26040)) {
                                    HSLFTextRunList.get(j).setFontSize(20.0);
                                }
                                // 设置字体样式为宋体
                                // String
                                // family=HSLFTextRunList.get(j).getFontFamily();
                                HSLFTextRunList.get(j).setFontFamily("宋体");
                            }
                        }

                    }
                    // 创建BufferedImage对象，图像的尺寸为原来的每页的尺寸*倍数
                    BufferedImage oneBufferedImage = new BufferedImage(onePPTPageSize.width * multiple,
                            onePPTPageSize.height * multiple, BufferedImage.TYPE_INT_RGB);
                    Graphics2D oneGraphics2D = oneBufferedImage.createGraphics();
                    // 设置转换后的图片背景色为白色
                    oneGraphics2D.setPaint(Color.white);
                    oneGraphics2D.scale(multiple, multiple);// 将图片放大multiple倍
                    oneGraphics2D
                            .fill(new Rectangle2D.Float(0, 0, onePPTPageSize.width * multiple, onePPTPageSize.height * multiple));
                    slide.draw(oneGraphics2D);
                    // 把BufferedImage转换为Base64
                    String base64 = BufferedImageToBase64(oneBufferedImage,"jpg");
                    if(null != base64 && !"".equals(base64.trim()))
                        base64List.add(base64);
                }

                String html = createPPTHtml(base64List);
                if(html != null){
                    try(FileOutputStream fos = new FileOutputStream(htmlSavePath + File.separator + htmlName)){
                        fos.write(html.getBytes("gbk"));
                    }
                }
                return htmlName;
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        return null;
    }

    /**
     * 将PPTX 文件转换成image
     *
     * @param sourcePath PPT文件路径
     * @param htmlSavePath 生成的html存储的文件夹路径
     * @param multiple 图片放大的倍数，提高清晰度
     * @return String html文件的名称
     */
    public String PPT2007toHtml(String sourcePath, String htmlSavePath, int multiple) {
        List<String> base64List = new ArrayList<>();
        String htmlName = SHA256.getSHA(sourcePath) + ".html";
        File fileDir = new File(htmlSavePath + File.separator + htmlName);
        // 如果已经存在，直接返回
        if (fileDir.exists()) {
            return htmlName;
        }
        try(FileInputStream fis = new FileInputStream(sourcePath)) {
            try(XMLSlideShow xss = new XMLSlideShow (fis)) {
                // 获取PPT每页的大小（宽和高度）
                Dimension onePPTPageSize = xss.getPageSize();
                // 获得PPT文件中的所有的PPT页面（获得每一张幻灯片）,并转为一张张的播放片
                List<XSLFSlide> slideList = xss.getSlides();
                // 设置字体为宋体，防止中文乱码
                for (XSLFSlide slide:slideList) {
                    List<XSLFShape> shapes = slide.getShapes();
                    for (XSLFShape shape : shapes) {
                        if (shape instanceof XSLFTextShape) {
                            XSLFTextShape sh = (XSLFTextShape) shape;
                            List<XSLFTextParagraph> textParagraphs = sh.getTextParagraphs();
                            for (XSLFTextParagraph xslfTextParagraph : textParagraphs) {
                                List<XSLFTextRun> textRuns = xslfTextParagraph.getTextRuns();
                                for (XSLFTextRun xslfTextRun : textRuns) {
                                    xslfTextRun.setFontFamily("宋体");
                                }
                            }
                        }
                    }
                    // 创建BufferedImage对象，图像的尺寸为原来的每页的尺寸*倍数
                    BufferedImage oneBufferedImage = new BufferedImage(onePPTPageSize.width * multiple,
                            onePPTPageSize.height * multiple, BufferedImage.TYPE_INT_RGB);
                    Graphics2D oneGraphics2D = oneBufferedImage.createGraphics();
                    // 设置转换后的图片背景色为白色
                    oneGraphics2D.setPaint(Color.white);
                    oneGraphics2D.scale(multiple, multiple);// 将图片放大multiple倍
                    oneGraphics2D
                            .fill(new Rectangle2D.Float(0, 0, onePPTPageSize.width * multiple, onePPTPageSize.height * multiple));
                    slide.draw(oneGraphics2D);
                    // 把BufferedImage转换为Base64
                    String base64 = BufferedImageToBase64(oneBufferedImage,"jpg");
                    if(null != base64 && !"".equals(base64.trim()))
                        base64List.add(base64);
                }

                String html = createPPTHtml(base64List);
                if(html != null){
                    try(FileOutputStream fos = new FileOutputStream(htmlSavePath + File.separator + htmlName)){
                        fos.write(html.getBytes("gbk"));
                    }
                }
                return htmlName;
            }

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        return null;
    }

    /**
     * 把BufferedImage转换为Base64形式
     *
     * @param bufferedImage PPT文件路径
     * @param format 图片的格式 (例如png,jpg)
     * @return
     */
    private String BufferedImageToBase64(BufferedImage bufferedImage,String format){
        String base64 = "";
        try(ByteArrayOutputStream baos = new ByteArrayOutputStream()){
            ImageIO.write(bufferedImage,format,baos);
            base64 = "data:image/" + format + ";base64," + DatatypeConverter.printBase64Binary(baos.toByteArray());
        }catch (IOException e) {
            return base64;
        }
        return base64;
    }

    /**
     * 把BufferedImage转换存入文件夹
     *
     * @param bufferedImage PPT文件路径
     * @param format 保存的图片的格式 (例如png,jpg)
     * @param imageFileDir 图片保存的路径
     * @return String 图片保存的路径
     */
    private String saveBufferedImage(BufferedImage bufferedImage,String format,String imageFileDir){
        // 存储图像的名称
        String imgName = UUID.randomUUID().toString() + "." + format;
        String imgSavePath;
        try(FileOutputStream fos = new FileOutputStream((imgSavePath = imageFileDir + File.separator + imgName))) {
            // 转换后的图片文件保存的指定的目录中
            ImageIO.write(bufferedImage, format, fos);
            return imgSavePath;
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }

    /**
     * 创建一个PPT的html
     *
     * @param base64List 图片的base64
     * @return String 生成的html
     */
    private String createPPTHtml(List<String> base64List){
        if(base64List == null || base64List.size() == 0)
            return null;
        StringBuffer sb = new StringBuffer();
        sb.append("<!DOCTYPE html><html lang=\"zh-CN\"><head><meta charset=\"UTF-8\">")
                .append("<meta name=\"viewport\" content=\"width=device-width, initial-scale=1,")
                .append("minimum-scale=1, maximum-scale=1\"><title>ppt</title>")
                .append("<link type=\"text/css\" rel=\"stylesheet\" href=\"ppt/swiper.min.css\" />")
                .append("<link type=\"text/css\" rel=\"stylesheet\" href=\"ppt/ppt.css\" />")
                .append("</head><body><div class=\"menu\">");
        StringBuffer menuImg = new StringBuffer();
        StringBuffer swiperSlide = new StringBuffer();
        for(String base64:base64List){
            menuImg.append("<div class=\"menu-img")
                    .append((base64List.indexOf(base64) == 0) ? " active" : "")
                    .append("\"><img src=\"")
                    .append(base64)
                    .append("\" width=\"100%\"></div>");
            swiperSlide.append("<div class=\"swiper-slide\"><img src=\"")
                    .append(base64)
                    .append("\" alt=\"\"></div>");
        }
        sb.append(menuImg.toString());
        sb.append("</div><div class=\"wrap\"><div class=\"swiper-container pptSwiper\"><div class=\"swiper-wrapper\">")
                .append(swiperSlide.toString())
                .append("</div></div><div class=\"slideshow\">Auto Play&nbsp;&nbsp;")
                .append("<img id=\"play_icon\" src=\"ppt/play.png\"></div></div>")
                .append("<script language=\"javascript\" src=\"js/jquery.min.js\"></script>")
                .append("<script language=\"javascript\" src=\"ppt/swiper.min.js\"></script>")
                .append("<script language=\"javascript\" src=\"ppt/ppt.js\"></script>")
                .append("</body></html>");
        return sb.toString();
    }

    public static void main(String[] args) {
        PPTToHtml pptToHtml = new PPTToHtml();
        String result =
                pptToHtml.PPT2007toHtml(
                        "F:\\LP-T07-05 Sample Completion.pptx", "F:\\测试\\", 8);
        System.out.println(result);
    }
}