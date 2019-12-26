package html;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.formula.ConditionalFormattingEvaluator;
import org.apache.poi.ss.formula.EvaluationConditionalFormatRule;
import org.apache.poi.ss.formula.WorkbookEvaluatorProvider;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class ExcelToHtml {
    static FormulaEvaluator evaluator;

    private static final Pattern pattern = Pattern.compile("^IFERROR\\((.*),{1}(.*)\\)$");

    private static CellStyle conditionalFormulaStyle;

    private static boolean isXSSF = true;

    private static Workbook workbook;

    /**
     * Excel转换为html
     *
     * @param sourcePath
     *          excel文件路径
     * @param savePath
     *          存储路径
     * @param saveName
     *          存储名称
     * @return
     */
    static String conversion(String sourcePath,String savePath,String saveName)
            throws FileNotFoundException,IOException{
        String path = null;
        String html = "";
        try(InputStream is = new FileInputStream(new File(sourcePath))) {
            Workbook wb = WorkbookFactory.create(is);
            workbook = wb;
            evaluator = wb.getCreationHelper().createFormulaEvaluator();
            conditionalFormulaStyle = wb.createCellStyle();
            if (wb instanceof XSSFWorkbook) {
                XSSFWorkbook xWb = (XSSFWorkbook) wb;
                html = getExcelInfo(xWb,true);
            }else if(wb instanceof HSSFWorkbook){
                isXSSF = false;
                HSSFWorkbook hWb = (HSSFWorkbook) wb;
                html = getExcelInfo(hWb,true);
            }
            File file = new File(savePath + File.separator + saveName + ".html");
            if(file.exists()){
                return file.getPath();
            }
            try(OutputStream os = new FileOutputStream(file)){
                os.write(html.getBytes());
            }
        }
        return path;
    }

    /**
     * 获取Excel信息
     * @param wb
     * @param isWithStyle
     * @return
     */
    static String getExcelInfo(Workbook wb, boolean isWithStyle){
        workbook = wb;
        evaluator = wb.getCreationHelper().createFormulaEvaluator();
        conditionalFormulaStyle = wb.createCellStyle();
        if(wb instanceof HSSFWorkbook){
            isXSSF = false;
        }
        String domain = System.getProperty("BASF-DOMAIN");
        domain = domain == null ? "" : domain;
        StringBuffer sb = new StringBuffer();
        sb.append("<!DOCTYPE html><html><head>")
                .append("<script language=\"javascript\" src=\""+domain+"WEB-JSP/js/jquery.min.js\"></script>")
                .append("<script language=\"javascript\" src=\""+domain+"WEB-JSP/js/common.js\"></script>")
                .append("<link type=\"text/css\" rel=\"stylesheet\" href=\""+domain+"WEB-JSP/css/style.css\" />")
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

            Sheet sheet = wb.getSheetAt(i);

            // get conditional formatting in the sheet
            SheetConditionalFormatting formatting = sheet.getSheetConditionalFormatting();

            int formattingsNum = formatting.getNumConditionalFormattings();

            ConditionalFormatting[] conditionalFormattings = new ConditionalFormatting[0];
            if (formattingsNum != 0){
                conditionalFormattings = new ConditionalFormatting[formattingsNum];
                // get all conditional formatting
                for (int j = 0; j < formattingsNum; j++) {
                    ConditionalFormatting conditionalFormatting = formatting.getConditionalFormattingAt(j);
                    conditionalFormattings[j] = conditionalFormatting;
                }
            }

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

                    String stringValue = getCellValue(cell,conditionalFormattings);
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
    static Map<String, String>[] getRowSpanColSpanMap(Sheet sheet) {

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
    static String getCellValue(Cell cell,ConditionalFormatting[] conditionalFormattings) {
        String result = new String();
        switch (cell.getCellType()) {
            case NUMERIC:// 数字类型
                if (DateUtil.isCellDateFormatted(cell)) {// 处理日期格式、时间格式
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
                    // 单元格设置成常规
                    if ("General".equals(temp)) {
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
            case FORMULA:
                try{
                    CellValue cellValue = evaluator.evaluate(cell);
                    if (cellValue.getCellType() == CellType.STRING){
                        result = cellValue.getStringValue();
                    }
                    else if(cellValue.getCellType() == CellType.NUMERIC){
                        result = Double.toString(cellValue.getNumberValue());
                    }
                } catch (Exception e){
                    String formula = cell.getCellFormula();
                    Matcher matcher = pattern.matcher(formula);
                    if (matcher.find()){
                        String value = matcher.group();
                        String[] strs = value.split(",");
                        int length = 0;
                        if ((length = strs.length) > 0){
                            String errorValue = strs[length - 1];
                            if (Objects.nonNull(errorValue) && errorValue.length() > 1){
                                result = errorValue.substring(0,errorValue.length() - 1);
                                cell.setCellValue(result);
                                break;
                            }
                        }
                    }
                    result = "";
                }
                break;
            default:
                result = "";
                break;
        }

        // cell上匹配成功的条件格式规则
        List<EvaluationConditionalFormatRule> mathcerRules = getMatchingConditionalFormattingForCell(cell);

        for (EvaluationConditionalFormatRule ruleEvaluation:mathcerRules){
            ConditionalFormattingRule cFRule = ruleEvaluation.getRule();
            PatternFormatting patternFormatting;
            if ((patternFormatting = cFRule.getPatternFormatting()) != null){
                // 获取填充色
                Color color = patternFormatting.getFillBackgroundColorColor();
                // 暂时只处理了 PatternFormatting 的填充色，后续有需再添加去了
                addConditionStyle(cell,color);
            }
        }
        return result;
    }

    private static void addConditionStyle(Cell cell,Color color){
        if (cell == null || color == null)
            return;
        conditionalFormulaStyle = workbook.createCellStyle();
        if (isXSSF){
            ((XSSFCellStyle)conditionalFormulaStyle).setFillForegroundColor((XSSFColor) color);
        }
        else {
            conditionalFormulaStyle.setFillForegroundColor(((HSSFColor)color).getIndex());
        }
        // 符合条件需要添加指定样式
        cell.setCellStyle(conditionalFormulaStyle);
    }

    /**
     * 处理表格样式
     * @param wb
     * @param sheet
     * @param cell
     * @param sb
     */
    static void dealExcelStyle(Workbook wb, Sheet sheet, Cell cell, StringBuffer sb){

        CellStyle cellStyle = cell.getCellStyle();
        if (cellStyle != null) {
            short alignment = cellStyle.getAlignment().getCode();
            sb.append("align='" + convertAlignToHtml(alignment) + "' ");//单元格内容的水平对齐方式
            short verticalAlignment = cellStyle.getAlignment().getCode();
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
                BorderStyle border = cellStyle.getBorderBottom();
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
                BorderStyle borderStyle = cellStyle.getBorderBottom();
                sb.append( getBorderStyle(palette,0,borderStyle.getCode(),cellStyle.getTopBorderColor()));
                sb.append( getBorderStyle(palette,1,borderStyle.getCode(),cellStyle.getRightBorderColor()));
                sb.append( getBorderStyle(palette,3,borderStyle.getCode(),cellStyle.getLeftBorderColor()));
                sb.append( getBorderStyle(palette,2,borderStyle.getCode(),cellStyle.getBottomBorderColor()));
            }

            sb.append("' ");
        }
    }

    /**
     * 单元格内容的水平对齐方式
     * @param alignment
     * @return
     */
    static String convertAlignToHtml(short alignment) {
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
    static String convertVerticalAlignToHtml(short verticalAlignment) {

        String valign = "middle";
        if(verticalAlignment == VerticalAlignment.BOTTOM.getCode()){valign = "bottom";}
        else if(verticalAlignment == VerticalAlignment.CENTER.getCode()){valign = "center";}
        else if(verticalAlignment == VerticalAlignment.TOP.getCode()){valign = "top";}
        return valign;
    }

    static String convertToStardColor(HSSFColor hc) {

        StringBuffer sb = new StringBuffer("");
        if (hc != null) {
            if (HSSFColor.HSSFColorPredefined.AUTOMATIC.getIndex() == hc.getIndex()) {
                return null;
            }
            sb.append("#");
            for (int i = 0; i < hc.getTriplet().length; i++) {
                sb.append(fillWithZero(Integer.toHexString(hc.getTriplet()[i])));
            }
        }

        return sb.toString();
    }

    static String fillWithZero(String str) {
        if (str != null && str.length() < 2) {
            return "0" + str;
        }
        return str;
    }

    static String[] bordesr={"border-top:","border-right:","border-bottom:","border-left:"};
    static String[] borderStyles={"solid ","solid ","solid ","solid ","solid ","solid ","solid ","solid ","solid ","solid","solid","solid","solid","solid"};

    static String getBorderStyle(HSSFPalette palette , int b, short s, short t){

        if(s==0)return  bordesr[b]+borderStyles[s]+"#d0d7e5 1px;";
        String borderColorStr = convertToStardColor( palette.getColor(t));
        borderColorStr=borderColorStr==null|| borderColorStr.length()<1?"#000000":borderColorStr;
        return bordesr[b]+borderStyles[s]+borderColorStr+" 1px;";

    }

    static String getBorderStyle(int b,short s, XSSFColor xc){

        if(s==0)return  bordesr[b]+borderStyles[s]+"#d0d7e5 1px;";;
        if (xc != null && !"".equals(xc)) {
            String borderColorStr = xc.getARGBHex();//t.getARGBHex();
            borderColorStr=borderColorStr==null|| borderColorStr.length()<1?"#000000":borderColorStr.substring(2);
            return bordesr[b]+borderStyles[s]+borderColorStr+" 1px;";
        }

        return "";
    }

    private static void usage(String error) throws IOException {
        String msg = "Excel转HTML出现错误\n"
                + (error == null ? "" : "错误信息: " + error + "\n") +
                "参数选项:\n" +
                "-source <必填> Excel文件路径\n" +
                "-save <必填> 存储路径 \n" +
                "-name <必填> 存储的html名称\n";
        throw new IllegalArgumentException(msg);
    }

    // 获取某个cell所有的条件格式规则
    private static List<EvaluationConditionalFormatRule> getMatchingConditionalFormattingForCell(Cell cell){
        List<EvaluationConditionalFormatRule> rules = new ArrayList<>();

        // Workbook各种评估器的提供者
        WorkbookEvaluatorProvider workbookEvaluatorProvider =
                (WorkbookEvaluatorProvider)workbook.getCreationHelper().createFormulaEvaluator();

        // 条件格式评估器初始化
        ConditionalFormattingEvaluator conditionalFormattingEvaluator =
                new ConditionalFormattingEvaluator(workbook,workbookEvaluatorProvider);

        // 获取cell的 规则评估器
        List<EvaluationConditionalFormatRule> allCFRulesForCell = conditionalFormattingEvaluator.getConditionalFormattingForCell(cell);

        for (EvaluationConditionalFormatRule evalCFRule :allCFRulesForCell) {
            // 获取某个规则涵盖的所有cell matchingCells
            List<Cell> matchingCells = conditionalFormattingEvaluator.getMatchingCells(evalCFRule);
            // 如果 matchingCells 涵盖当前的cell，表示当前的cell含有此条规则，添加
            if (matchingCells.contains(cell)) rules.add(evalCFRule);
        }

        return rules;
    }

    public static void main(String[] args) {
        try {
            if (args == null || args.length == 0) {
                usage("没有传入必填的参数！");
                return;
            }
            String source = "";
            String save = "";
            String name = "";
            for(int i = 0; i < args.length / 2; ++i) {
                String key = args[i * 2];
                String value = i * 2 + 1 < args.length ? args[i * 2 + 1] : "";
                switch(key.hashCode()) {
                    case 386454152:
                        if ("-source".equals(key)) {
                            source = value;
                        }
                        break;
                    case 45081386:
                        if ("-save".equals(key)) {
                            save = value;
                        }
                        break;
                    case 44932152:
                        if ("-name".equals(key)) {
                            name = value;
                        }
                        break;
                    default:
                        break;
                }
            }
            if ("".equals(source)){
                usage("参数 -source 是必填项！");
                return;
            }
            if ("".equals(save)){
                usage("参数 -save 是必填项！");
                return;
            } if ("".equals(name)){
                usage("参数 -name 是必填项！");
                return;
            }
            if(!source.endsWith("xlsx") && !source.endsWith("xls")){
                usage("参数 -source 只能为xls和xlsx文件！");
                return;
            }
            conversion(source,save,name);
            System.out.println("successful");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}