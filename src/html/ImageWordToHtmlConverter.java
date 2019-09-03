package html;

import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.hwpf.usermodel.Picture;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

import java.util.Base64;

/**
 * 处理doc文件转换html时的图片显示
 */
public class ImageWordToHtmlConverter extends WordToHtmlConverter {

    public ImageWordToHtmlConverter(Document document) {
        super(document);
    }

    /**不适用PicturesManager的图片处理，图片转换为base64形式*/
    @Override
    protected void processImageWithoutPicturesManager(Element currentBlock,
                                                      boolean inlined, Picture picture) {
        Element imgNode = currentBlock.getOwnerDocument().createElement("img");
        StringBuilder sb = new StringBuilder();
        sb.append(Base64.getMimeEncoder().encodeToString(picture.getRawContent()));
        sb.insert(0, "data:" + picture.getMimeType() + ";base64,");
        imgNode.setAttribute("src", sb.toString());
        currentBlock.appendChild(imgNode);
    }

}