package html;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.apache.xmlbeans.impl.piccolo.io.FileFormatException;

import java.io.*;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

public class ExcelToHtml {
    /**
     * Excel转换为html
     *
     * @param sourcePath
     *          excel文件路径
     * @param sourceFileName
     *          excel文件名
     * @param savePath
     *          存储路径
     * @return
     * @return
     */
    public String excelToHtml(String sourcePath,String sourceFileName,String suffix,String savePath) throws IOException {
        if(".csv".equals(suffix)){
            try {
                String excelPath = csvToExcel(sourcePath,sourceFileName + suffix,savePath);
                File file = new File(excelPath);
                if(!file.exists())
                    throw new FileNotFoundException();
                // 重新赋值需要读取的文件的路径
                sourcePath = excelPath.substring(0,excelPath.lastIndexOf(File.separator)+1);
                sourceFileName = excelPath.substring(excelPath.lastIndexOf(File.separator)+1,excelPath.length());
                if(sourceFileName.lastIndexOf(".") < 0)
                    throw new FileFormatException("没有文件类型的文件，不支持预览");
                suffix = sourceFileName.substring(sourceFileName.lastIndexOf("."),sourceFileName.length());
                sourceFileName = sourceFileName.substring(0,sourceFileName.lastIndexOf("."));
            } catch (Exception e) {
                throw new EOFException(e.getMessage());
            }
            sourcePath = savePath;
        }

        // 获取源文件名称
        sourceFileName = sourceFileName + suffix;
        String excelName = SHA256.getSHA(sourceFileName);
        String htmlPath = savePath + File.separator + excelName+".html";
        File targetFile = new File(savePath+File.separator);
        if(!targetFile.exists()){
            targetFile.mkdirs();
        }
        // 读取源文件流
        InputStream is = null;
        // 写入目标文件流
        FileOutputStream outputStream = null;
        // 返回的html
        String htmlExcel = null;
        try {
            is = new FileInputStream(new File(sourcePath + File.separator +sourceFileName));
            Workbook wb = WorkbookFactory.create(is);
            // HSSFWorkbook（xls）和 XSSFWorkbook （xlsx）
            if (wb instanceof XSSFWorkbook) {
                XSSFWorkbook xWb = (XSSFWorkbook) wb;
                htmlExcel = getExcelInfo(xWb,true);
            }else if(wb instanceof HSSFWorkbook){
                HSSFWorkbook hWb = (HSSFWorkbook) wb;
                htmlExcel = getExcelInfo(hWb,true);
            }
            // 将内容存放到html中
            File html = new File(htmlPath);
            if(html.exists()){
                return excelName + ".html";
            }
            // 写入到指定的文件
            outputStream = new FileOutputStream(html);
            outputStream.write(htmlExcel.getBytes("gbk"));
        } catch (Exception e) {
            e.printStackTrace();
        }finally{
            try {
                if(is!=null){
                    is.close();
                }
                if(outputStream!=null){
                    outputStream.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return excelName + ".html";
    }

    /**
     * 获取Excel信息
     * @param wb
     * @param isWithStyle
     * @return
     */
    public String getExcelInfo(Workbook wb, boolean isWithStyle){
        StringBuffer sb = new StringBuffer();
        sb.append("<!DOCTYPE html><html><head>")
                .append("<script language=\"javascript\" src=\"js/jquery.min.js\"></script>")
                .append("<script language=\"javascript\" src=\"js/common.js\"></script>")
                .append("<link type=\"text/css\" rel=\"stylesheet\" href=\"css/style.css\" />")
                .append("</head><body>");

        sb.append("<div class=\"investment_f\">").append("<div class=\"investment_title\">");

        for(int i = 0; i < wb.getNumberOfSheets(); i++){//设置切换按钮
            String sheetName = wb.getSheetName(i);
            if(i==0){
                sb.append("<div class=\"on\">"+sheetName+"</div>");
            }else{
                sb.append("<div>"+sheetName+"</div>");
            }
        }
        sb.append("</div>").append("<div class=\"investment_con\">");
        for(int i = 0; i < wb.getNumberOfSheets(); i++){
            sb.append("<div class=\"investment_con_list\">");
            Sheet sheet = wb.getSheetAt(i);//获取每一个Sheet的内容
            sb.append("<div class='tab"+i+"'>");

            int lastRowNum = sheet.getLastRowNum();//获取最后一行的编号
            Map<String, String> map[] = getRowSpanColSpanMap(sheet);
            sb.append("<table style='border-collapse:collapse;' width='100%'>");
            Row row = null;
            Cell cell = null;

            for (int rowNum = sheet.getFirstRowNum(); rowNum <= lastRowNum; rowNum++) {//遍历获取每一行
                row = sheet.getRow(rowNum);
                if (row == null) {
                    sb.append("<tr><td >&nbsp;&nbsp;</td></tr>");
                    continue;
                }
                sb.append("<tr>");
                int lastColNum = row.getLastCellNum();//获取最后一列
                for (int colNum = 0; colNum < lastColNum; colNum++) {//遍历每一列
                    cell = row.getCell(colNum);
                    if (cell == null) {    //特殊情况 空白的单元格会返回null
                        sb.append("<td>&nbsp;</td>");
                        continue;
                    }

                    String stringValue = getCellValue(cell);
                    if (map[0].containsKey(rowNum + "," + colNum)) {
                        String pointString = map[0].get(rowNum + "," + colNum);
                        map[0].remove(rowNum + "," + colNum);
                        int bottomeRow = Integer.valueOf(pointString.split(",")[0]);
                        int bottomeCol = Integer.valueOf(pointString.split(",")[1]);
                        int rowSpan = bottomeRow - rowNum + 1;
                        int colSpan = bottomeCol - colNum + 1;
                        sb.append("<td rowspan= '" + rowSpan + "' colspan= '"+ colSpan + "' ");
                    } else if (map[1].containsKey(rowNum + "," + colNum)) {
                        map[1].remove(rowNum + "," + colNum);
                        continue;
                    } else {
                        sb.append("<td ");
                    }

                    //判断是否需要样式
                    if(isWithStyle){
                        dealExcelStyle(wb, sheet, cell, sb);//处理单元格样式
                    }

                    sb.append(">");
                    if (stringValue == null || "".equals(stringValue.trim())) {
                        sb.append("&nbsp;&nbsp;&nbsp;");
                    } else {
                        // 将ascii码为160的空格转换为html下的空格（ ）
                        sb.append(stringValue.replace(String.valueOf((char) 160),"&nbsp;&nbsp;"));
                    }
                    sb.append("</td>");
                }
                sb.append("</tr>");
            }

            sb.append("</table>");
            sb.append("</div>").append("</div>");
        }
        sb.append("</div></body></html>");
        return sb.toString();
    }

    /**
     * 获取合并单元格信息
     * @param sheet
     * @return
     */
    private Map<String, String>[] getRowSpanColSpanMap(Sheet sheet) {

        Map<String, String> map0 = new HashMap<String, String>();
        Map<String, String> map1 = new HashMap<String, String>();
        int mergedNum = sheet.getNumMergedRegions();
        CellRangeAddress range = null;
        for (int i = 0; i < mergedNum; i++) {
            range = sheet.getMergedRegion(i);
            int topRow = range.getFirstRow();
            int topCol = range.getFirstColumn();
            int bottomRow = range.getLastRow();
            int bottomCol = range.getLastColumn();
            map0.put(topRow + "," + topCol, bottomRow + "," + bottomCol);
            // System.out.println(topRow + "," + topCol + "," + bottomRow + "," + bottomCol);
            int tempRow = topRow;
            while (tempRow <= bottomRow) {
                int tempCol = topCol;
                while (tempCol <= bottomCol) {
                    map1.put(tempRow + "," + tempCol, "");
                    tempCol++;
                }
                tempRow++;
            }
            map1.remove(topRow + "," + topCol);
        }
        Map[] map = { map0, map1 };
        return map;
    }

    /**
     * 获取表格单元格Cell内容
     * @param cell
     * @return
     */
    private String getCellValue(Cell cell) {

        String result = new String();
        switch (cell.getCellTypeEnum()) {
            case NUMERIC:// 数字类型
                if (HSSFDateUtil.isCellDateFormatted(cell)) {// 处理日期格式、时间格式
                    SimpleDateFormat sdf = null;
                    if (cell.getCellStyle().getDataFormat() == HSSFDataFormat.getBuiltinFormat("h:mm")) {
                        sdf = new SimpleDateFormat("HH:mm");
                    } else {// 日期
                        sdf = new SimpleDateFormat("yyyy-MM-dd");
                    }
                    Date date = cell.getDateCellValue();
                    result = sdf.format(date);
                } else if (cell.getCellStyle().getDataFormat() == 58) {
                    // 处理自定义日期格式：m月d日(通过判断单元格的格式id解决，id的值是58)
                    SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                    double value = cell.getNumericCellValue();
                    Date date = DateUtil
                            .getJavaDate(value);
                    result = sdf.format(date);
                } else {
                    double value = cell.getNumericCellValue();
                    CellStyle style = cell.getCellStyle();
                    DecimalFormat format = new DecimalFormat();
                    String temp = style.getDataFormatString();
                    // System.out.println(temp);
                    // 单元格设置成常规
                    if ("General".equals(temp)) {
//                    	if (temp.equals("General")) {
                        format.applyPattern("#");
                    }
                    result = format.format(value);
                }
                break;
            case STRING:// String类型
                result = cell.getRichStringCellValue().toString();
                break;
            case BLANK:
                result = "";
                break;
            default:
                result = "";
                break;
        }
        return result;
    }

    /**
     * 处理表格样式
     * @param wb
     * @param sheet
     * @param cell
     * @param sb
     */
    private void dealExcelStyle(Workbook wb, Sheet sheet, Cell cell, StringBuffer sb){

        CellStyle cellStyle = cell.getCellStyle();
        if (cellStyle != null) {
            short alignment = cellStyle.getAlignmentEnum().getCode();
            sb.append("align='" + convertAlignToHtml(alignment) + "' ");//单元格内容的水平对齐方式
            short verticalAlignment = cellStyle.getVerticalAlignmentEnum().getCode();
            sb.append("valign='"+ convertVerticalAlignToHtml(verticalAlignment)+ "' ");//单元格中内容的垂直排列方式

            if (wb instanceof XSSFWorkbook) {

                XSSFFont xf = ((XSSFCellStyle) cellStyle).getFont();
                sb.append("style='");
                if(xf.getBold())
                    sb.append("font-weight:bold;"); // 字体加粗
                sb.append("font-size: " + xf.getFontHeight() / 2 + "%;"); // 字体大小
                int columnWidth = sheet.getColumnWidth(cell.getColumnIndex()) ;
                sb.append("width:" + columnWidth + "px;");

                XSSFColor xc = xf.getXSSFColor();
                if (xc != null && !"".equals(xc)) {
                    sb.append("color:#" + xc.getARGBHex().substring(2) + ";"); // 字体颜色
                }
                XSSFColor bgColor = (XSSFColor) cellStyle.getFillForegroundColorColor();
                if (bgColor != null && !"".equals(bgColor)) {
                    sb.append("background-color:#" + bgColor.getARGBHex().substring(2) + ";"); // 背景颜色
                }
                BorderStyle border = cellStyle.getBorderBottomEnum();
                sb.append(getBorderStyle(0,border.getCode(), ((XSSFCellStyle) cellStyle).getTopBorderXSSFColor()));
                sb.append(getBorderStyle(1,border.getCode(), ((XSSFCellStyle) cellStyle).getRightBorderXSSFColor()));
                sb.append(getBorderStyle(2,border.getCode(), ((XSSFCellStyle) cellStyle).getBottomBorderXSSFColor()));
                sb.append(getBorderStyle(3,border.getCode(), ((XSSFCellStyle) cellStyle).getLeftBorderXSSFColor()));

            }else if(wb instanceof HSSFWorkbook){

                HSSFFont hf = ((HSSFCellStyle) cellStyle).getFont(wb);
                short fontColor = hf.getColor();
                sb.append("style='");
                HSSFPalette palette = ((HSSFWorkbook) wb).getCustomPalette(); // 类HSSFPalette用于求的颜色的国际标准形式
                HSSFColor hc = palette.getColor(fontColor);
                if(hf.getBold())
                    sb.append("font-weight:bold;"); // 字体加粗
                sb.append("font-size: " + hf.getFontHeight() / 2 + "%;"); // 字体大小
                String fontColorStr = convertToStardColor(hc);
                if (fontColorStr != null && !"".equals(fontColorStr.trim())) {
                    sb.append("color:" + fontColorStr + ";"); // 字体颜色
                }
                int columnWidth = sheet.getColumnWidth(cell.getColumnIndex()) ;
                sb.append("width:" + columnWidth + "px;");
                short bgColor = cellStyle.getFillForegroundColor();
                hc = palette.getColor(bgColor);
                String bgColorStr = convertToStardColor(hc);
                if (bgColorStr != null && !"".equals(bgColorStr.trim())) {
                    sb.append("background-color:" + bgColorStr + ";"); // 背景颜色
                }
                sb.append( getBorderStyle(palette,0,cellStyle.getBorderTop(),cellStyle.getTopBorderColor()));
                sb.append( getBorderStyle(palette,1,cellStyle.getBorderRight(),cellStyle.getRightBorderColor()));
                sb.append( getBorderStyle(palette,3,cellStyle.getBorderLeft(),cellStyle.getLeftBorderColor()));
                sb.append( getBorderStyle(palette,2,cellStyle.getBorderBottom(),cellStyle.getBottomBorderColor()));
            }

            sb.append("' ");
        }
    }

    /**
     * 单元格内容的水平对齐方式
     * @param alignment
     * @return
     */
    private String convertAlignToHtml(short alignment) {
        String align = "left";
        if(alignment == HorizontalAlignment.LEFT.getCode()){ align = "left";}
        else if(alignment == HorizontalAlignment.CENTER.getCode()){ align = "center";}
        else if(alignment == HorizontalAlignment.RIGHT.getCode()){ align = "right";}
        return align;
    }

    /**
     * 单元格中内容的垂直排列方式
     * @param verticalAlignment
     * @return
     */
    private String convertVerticalAlignToHtml(short verticalAlignment) {

        String valign = "middle";
        if(verticalAlignment == VerticalAlignment.BOTTOM.getCode()){valign = "bottom";}
        else if(verticalAlignment == VerticalAlignment.CENTER.getCode()){valign = "center";}
        else if(verticalAlignment == VerticalAlignment.TOP.getCode()){valign = "top";}
        return valign;
    }

    private String convertToStardColor(HSSFColor hc) {

        StringBuffer sb = new StringBuffer("");
        if (hc != null) {
            if (HSSFColor.AUTOMATIC.index == hc.getIndex()) {
                return null;
            }
            sb.append("#");
            for (int i = 0; i < hc.getTriplet().length; i++) {
                sb.append(fillWithZero(Integer.toHexString(hc.getTriplet()[i])));
            }
        }

        return sb.toString();
    }

    private String fillWithZero(String str) {
        if (str != null && str.length() < 2) {
            return "0" + str;
        }
        return str;
    }

    String[] bordesr={"border-top:","border-right:","border-bottom:","border-left:"};
    String[] borderStyles={"solid ","solid ","solid ","solid ","solid ","solid ","solid ","solid ","solid ","solid","solid","solid","solid","solid"};

    private String getBorderStyle(HSSFPalette palette , int b, short s, short t){

        if(s==0)return  bordesr[b]+borderStyles[s]+"#d0d7e5 1px;";
        String borderColorStr = convertToStardColor( palette.getColor(t));
        borderColorStr=borderColorStr==null|| borderColorStr.length()<1?"#000000":borderColorStr;
        return bordesr[b]+borderStyles[s]+borderColorStr+" 1px;";

    }

    private String getBorderStyle(int b,short s, XSSFColor xc){

        if(s==0)return  bordesr[b]+borderStyles[s]+"#d0d7e5 1px;";;
        if (xc != null && !"".equals(xc)) {
            String borderColorStr = xc.getARGBHex();//t.getARGBHex();
            borderColorStr=borderColorStr==null|| borderColorStr.length()<1?"#000000":borderColorStr.substring(2);
            return bordesr[b]+borderStyles[s]+borderColorStr+" 1px;";
        }

        return "";
    }

    /**
     * 把csv转换为xlsx存储
     *
     * @param sourcePath csv文件路径
     * @param sourceFileName 源文件名
     * @param savaPath 存储路径
     * @return String 转换完成的xlsx的存储路径
     * */
    private String csvToExcel(String sourcePath,String sourceFileName,String savaPath) throws Exception{
        ArrayList arList;
        ArrayList al;
        int i = 0;
        String thisLine;
        String excelSavePath = savaPath + File.separator + SHA256.getSHA(sourceFileName) + ".xlsx";
        try(InputStreamReader fr = new InputStreamReader(
                new FileInputStream(sourcePath + File.separator + sourceFileName),"utf-8");
            BufferedReader br = new BufferedReader(fr)){
            arList = new ArrayList();
            // 获取csv文件的行，存储在arList当中
            while ((thisLine = br.readLine()) != null) {
                al = new ArrayList();
                // 获取每一列
                String strar[] = thisLine.split(",");
                for(int j=0;j<strar.length;j++) {
                    al.add(strar[j]);
                }
                arList.add(al);
                i++;
            }

            // 创建工作表
            XSSFWorkbook hwb = new XSSFWorkbook();
            XSSFSheet sheet = hwb.createSheet("sheet1");
            // 读取csv的行
            for(int k=0;k<arList.size();k++) {
                // 读取列
                ArrayList ardata = (ArrayList)arList.get(k);
                XSSFRow row = sheet.createRow((short) 0+k);
                // 转换写入工作表行
                for(int p=0;p<ardata.size();p++) {
                    XSSFCell cell = row.createCell((short) p);
                    String data = ardata.get(p).toString();
                    if(data.startsWith("=")){
                        cell.setCellType(CellType.STRING);
                        data=data.replaceAll("\"", "");
                        data=data.replaceAll("=", "");
                        cell.setCellValue(data);
                    }else if(data.startsWith("\"")){
                        data=data.replaceAll("\"", "");
                        cell.setCellType(CellType.STRING);
                        cell.setCellValue(data);
                    }else{
                        data=data.replaceAll("\"", "");
                        cell.setCellType(CellType.NUMERIC);
                        cell.setCellValue(data);
                    }
                }
            }

            try(FileOutputStream fos = new FileOutputStream(excelSavePath)){
                hwb.write(fos);
            }
        }
        catch (UnsupportedEncodingException e) {
            throw new UnsupportedEncodingException("不支持的csv编码格式");
        } catch (FileNotFoundException e) {
            throw new FileNotFoundException();
        } catch (IOException e) {
            throw new IOException();
        }
        return excelSavePath;
    }

    public static void main(String[] args) {
        ExcelToHtml excelToHtml = new ExcelToHtml();
        try {
            System.out.println(excelToHtml.excelToHtml("F:\\","xlsx",".xlsx","F:\\测试"));
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}