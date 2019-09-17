package export;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.*;
import java.util.*;
import java.util.function.Predicate;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

/**
 * 使用一个已经存在的Excel作为模板，可以对当前的模板Excel进行修改操作，
 * 然后重新输出为流，或者存入文件系统当中。
 *
 * ExcelTemplate可以使用当前Excel中已经存在的行作为模板，插入新的行。
 * 比如说excel中的某些行，具有复杂的合并单元格，背景色等。但是我们现在
 * 需要以这些复杂的行作为模板，动态的在Excel中插入这种复杂的行，可以如下操作：
 * // 加载Excel模板
 * ExcelTemplate excel = new ExcelTemplate("F:\\加班表.xlsx");
 * // 验证是否可用
 * if(excel.examine()){
 *     Map<Integer, LinkedList<String>> areaValue = new LinkedHashMap<>();
 *     // 添加填充的数据
 *     LinkedList<String> array1 = new LinkedList<>();
 *     array1.add(Integer.toString(1));
 *     array1.add("123456");
 *     array1.add("张三");
 *     array1.add("2019/9/10");
 *     array1.add("2019/9/10");
 *     array1.add("2019/9/10");
 *     array1.add("2019/9/10");
 *     array1.add("项目加班");
 *     areaValue.put(1,array1);
 *     LinkedList<String> array2 = new LinkedList<>();
 *     array2.add(Integer.toString(1));
 *     array2.add("123456");
 *     array2.add("李四");
 *     array2.add("2019/9/10");
 *     array2.add("2019/9/10");
 *     array2.add("2019/9/10");
 *     array2.add("2019/9/10");
 *     array2.add("项目加班");
 *     areaValue.put(2,array2);
 *     try {
 *         excel.addRowByExist(16,16,17,areaValue,true);
 *         excel.save("F:\\测试\\poi.xlsx");
 *     } catch (InvalidFormatException e) {
 *         e.printStackTrace();
 *     } catch (IOException e) {
 *         e.printStackTrace();
 *     }
 * }
 *
 * 还有一种可能就是，模板中有些单元格的值需要动态的替换，可以使用
 * ${key} 来表示这个值是需要动态替换的，比如说有个单元格的值是要填
 * 写表格创建人，可以标记为 ${创建人}，然后如下调用：
 * // 加载Excel模板
 * ExcelTemplate excel = new ExcelTemplate("F:\\加班表.xlsx");
 * // 验证是否可用
 * if(excel.examine()){
 *     try {
 *         Map<String,String> map = new HashMap<>();
 *         map.put("创建人","张三");
 *         map.put("日期",new SimpleDateFormat("yyyy-MM-dd").format(new Date()));
 *         System.out.println("修改数量：" + excel.fillVariable(map));
 *         excel.save("F:\\测试\\poi.xlsx");
 *     } catch (InvalidFormatException e) {
 *         e.printStackTrace();
 *     } catch (IOException e) {
 *         e.printStackTrace();
 *     }
 * }
 * */
public class ExcelTemplate {
    private String path;

    private Workbook workbook;

    private Sheet[] sheets;

    private Sheet sheet;

    private Throwable ex;

    private List<Cell> cellList = null;

    /**
     * 通过模板Excel的路径初始化
     * */
    public ExcelTemplate(String path) {
        this.path = path;
        init();
    }

    private void init(){
        File file = new File(path);
        try (InputStream is = new FileInputStream(file)){
            workbook = WorkbookFactory.create(is);
            sheets = new Sheet[workbook.getNumberOfSheets()];
            for(int i = 0;i < sheets.length;i++){
                sheets[i] = workbook.getSheetAt(i);
            }
            if(sheets.length > 0)
                sheet = sheets[0];
        } catch (InvalidFormatException e) {
            ex = e;
        } catch (IOException e) {
            ex = e;
        }
    }

    private boolean initSheet(int sheetNo){
        if(sheetNo < 0 || sheetNo > sheets.length - 1)
            return false;
        sheet = sheets[sheetNo];
        return true;
    }

    /**
     * 验证模板是否可用
     * @return true-可用 false-不可用
     * */
    public boolean examine(){
        if(ex == null
                && workbook != null
                && sheets != null
                && sheets.length > 0
                && sheet != null)
            return true;
        return false;
    }

    private boolean examineSheetRow(int index){
        if(index < 0 || index > sheet.getLastRowNum())
            return false;
        return true;
    }

    /**
     * 使用一个已经存在的row作为模板，
     * 从sheet[0]的toRowNum行开始插入这个row模板的副本,
     * 并且使用areaValue从左至右，从上至下的替换掉
     * row区域中值为 ${} 的单元格的值
     *
     * @param fromRowIndex 模板行的索引
     * @param toRowIndex 开始插入的row索引
     * @param areaValues 替换模板row区域的${}值
     * @return int 插入的行数量
     * @throws IOException
     * @throws InvalidFormatException
     * */
    public int addRowByExist(int fromRowIndex, int toRowIndex,
                             LinkedHashMap<Integer,LinkedList<String>> areaValues)
            throws IOException, InvalidFormatException {
        return addRowByExist(0,fromRowIndex,fromRowIndex,toRowIndex,areaValues,true);
    }

    /**
     * 使用一个已经存在的row作为模板，
     * 从sheet[sheetNo]的toRowNum行开始插入这个row模板的副本
     *
     * @param sheetNo 需要操作的Sheet的编号
     * @param fromRowStartIndex 模板row区域的开始索引
     * @param fromRowEndIndex 模板row区域的结束索引
     * @param toRowIndex 开始插入的row索引值
     * @param copyNum 复制的数量
     * @param delRowTemp 是否删除模板row区域
     * @return int 插入的行数量
     * @throws IOException
     * @throws InvalidFormatException
     * */
    public int addRowByExist(int sheetNo,int fromRowStartIndex, int fromRowEndIndex,int toRowIndex, int copyNum,boolean delRowTemp)
            throws IOException, InvalidFormatException {
        LinkedHashMap<Integer, LinkedList<String>> map = new LinkedHashMap<>();
        for(int i = 1;i <= copyNum;i++){
            map.put(i,new LinkedList<>());
        }
        return addRowByExist(sheetNo,fromRowStartIndex,fromRowEndIndex,toRowIndex,map,delRowTemp);
    }

    /**
     * 使用一个已经存在的row作为模板，
     * 从sheet[0]的toRowNum行开始插入这个row模板的副本,
     * 并且使用areaValue从左至右，从上至下的替换掉
     * row区域中值为 ${} 的单元格的值
     *
     * @param fromRowStartIndex 模板row区域的开始索引
     * @param fromRowEndIndex 模板row区域的结束索引
     * @param toRowIndex 开始插入的row索引
     * @param areaValues 替换模板row区域的${}值
     * @param delRowTemp 是否删除模板row区域
     * @return int 插入的行数量
     * @throws IOException
     * @throws InvalidFormatException
     * */
    public int addRowByExist(int fromRowStartIndex, int fromRowEndIndex,int toRowIndex,
                             LinkedHashMap<Integer,LinkedList<String>> areaValues, boolean delRowTemp)
            throws IOException, InvalidFormatException {
        return addRowByExist(0,fromRowStartIndex,fromRowEndIndex,toRowIndex,areaValues,delRowTemp);
    }

    /**
     * 使用一个已经存在的row作为模板，
     * 从sheet[sheetNo]的toRowNum行开始插入这个row模板的副本,
     * 并且使用areaValue从左至右，从上至下的替换掉
     * row区域中值为 ${} 的单元格的值
     *
     * @param sheetNo 需要操作的Sheet的编号
     * @param fromRowIndex 模板行的索引
     * @param toRowIndex 开始插入的row索引
     * @param areaValues 替换模板row区域的${}值
     * @return int 插入的行数量
     * @throws IOException
     * @throws InvalidFormatException
     * */
    public int addRowByExist(int sheetNo,int fromRowIndex, int toRowIndex,
                             LinkedHashMap<Integer,LinkedList<String>> areaValues)
            throws IOException, InvalidFormatException {
        return addRowByExist(sheetNo,fromRowIndex,fromRowIndex,toRowIndex,areaValues,true);
    }

    /**
     * 使用一个已经存在的行区域作为模板，
     * 从sheet的toRowNum行开始插入这段行区域,
     * areaValue会从左至右，从上至下的替换掉row区域
     * 中值为 ${} 的单元格的值
     *
     * @param sheetNo 需要操作的Sheet的编号
     * @param fromRowStartIndex 模板row区域的开始索引
     * @param fromRowEndIndex 模板row区域的结束索引
     * @param toRowIndex 开始插入的row索引
     * @param areaValues 替换模板row区域的${}值
     * @param delRowTemp 是否删除模板row区域
     * @return int 插入的行数量
     * @throws IOException
     * @throws InvalidFormatException
     * */
    public int addRowByExist(int sheetNo,int fromRowStartIndex, int fromRowEndIndex,int toRowIndex,
                             LinkedHashMap<Integer,LinkedList<String>> areaValues, boolean delRowTemp)
            throws InvalidFormatException, IOException {
        exception();
        if(!examine()
                || !initSheet(sheetNo)
                || !examineSheetRow(fromRowStartIndex)
                || !examineSheetRow(fromRowEndIndex)
                || !examineSheetRow(toRowIndex)
                || fromRowStartIndex > fromRowEndIndex)
            return 0;
        int areaNum;List<Row> rows = new ArrayList<>();
        if(areaValues != null){
            int n = 0,f = areaValues.size() * (areaNum = (fromRowEndIndex - fromRowStartIndex + 1));
            // 在插入前腾出空间，避免新插入的行覆盖原有的行
            shiftAndCreateRows(sheetNo,toRowIndex,f);
            // 读取需要插入的数据
            for (Integer key:areaValues.keySet()){
                List<Row> temp = new LinkedList<>();
                // 插入行
                for(int i = sheet.getFirstRowNum();i < areaNum;i++){
                    int num = areaNum * n + i;
                    Row toRow = sheet.getRow(toRowIndex + num);
                    Row row;
                    if(toRowIndex >= fromRowEndIndex)
                        row = copyRow(sheetNo,sheet.getRow(fromRowStartIndex + i),sheetNo,toRow,true);
                    else
                        row = copyRow(sheetNo,sheet.getRow(fromRowStartIndex + i + f),sheetNo,toRow,true);
                    temp.add(row);
                }
                // 使用传入的值覆盖${}
                replaceMark(temp,areaValues.get(key));
                rows.addAll(temp);
                n++;
            }
            if(delRowTemp){
                for(int i = fromRowStartIndex;i <= fromRowEndIndex;i++){
                    Row row;
                    if(toRowIndex >= fromRowEndIndex && (row = sheet.getRow(i)) != null)
                        sheet.removeRow(row);
                    else if((row = sheet.getRow(i + f)) != null)
                        sheet.removeRow(row);
                }
            }
        }
        return rows.size();
    }

    /**
     * 填充Excel当中的变量
     *
     * @param fillValues 填充的值
     * @return int 受影响的变量数量
     * @throws IOException
     * @throws InvalidFormatException
     **/
    public int fillVariable(Map<String,String> fillValues) throws IOException, InvalidFormatException {
        return fillVariable(0,fillValues);
    }


    /**
     * 填充Excel当中的变量
     *
     * @param sheetNo 需要操作的Sheet的编号
     * @param fillValues 填充的值
     * @return int 受影响的变量数量
     * @throws IOException
     * @throws InvalidFormatException
     **/
    public int fillVariable(int sheetNo,Map<String,String> fillValues) throws IOException, InvalidFormatException {
        exception();
        if(!examine()
                || sheetNo < 0
                || sheetNo > sheets.length - 1
                || fillValues == null
                || fillValues.size() == 0)
            return 0;
        // 验证${}格式
        final Pattern pattern = Pattern.compile("(\\$\\{[^\\}]+})");
        // 把所有的${}按Cell分类，也就是说如果一个Cell中存在两个${}，
        // 这两个变量的Cell应该一样
        Map<Cell,Map<String,String>> cellVal = new HashMap<>();
        List<Integer> ns = new ArrayList<>();
        ns.add(0);
        fillValues.forEach((k,v) ->{
            // 找到变量所在的单元格
            Cell cell = find(sheetNo,s -> {
                if(s == null || "".equals(s))
                    return false;
                Matcher matcher = pattern.matcher(s);
                while(matcher.find()){
                    String variable = matcher.group(1);
                    if(variable != null
                            && formatParamCode(variable).equals(k.trim()))
                        return true;
                }
                return false;
            }).stream().findFirst().orElse(null);
            if(cell != null){
                Map<String,String> cellValMap = cellVal.get(cell);
                if(cellValMap == null)
                    cellValMap = new HashMap<>();
                cellValMap.put(k,v);
                cellVal.put(cell,cellValMap);
                ns.replaceAll(n -> n + 1);
            }
        });
        cellVal.forEach((k,v) -> {
            String cellValue = k.getStringCellValue();
            k.setCellValue(composeMessage(cellValue,v));
        });
        return ns.get(0);
    }

    /**
     * 根据断言predicate查找sheet当中符合条件的cell
     *
     * @param sheetNo 需要操作的Sheet的编号
     * @param predicate 筛选的断言
     * @return List<Cell> 符合条件的Cell
     * */
    private List<Cell> find(int sheetNo,Predicate<String> predicate){
        Objects.requireNonNull(predicate);
        initCellList(sheetNo);
        return cellList.stream()
                .map(c -> {
                    if(c != null && c.getCellTypeEnum() == CellType.STRING)
                        return c.getStringCellValue();
                    return null;
                })// Cell流转换为String流
                .filter(predicate)
                .map(s -> cellList.stream().filter(c -> {
                    if(c != null && c.getCellTypeEnum() == CellType.STRING
                            && s.equals(c.getStringCellValue()))
                        return true;
                    return false;
                }).findFirst().orElse(null))// String流重新转换位Cell流
                .filter(c -> c != null)
                .collect(Collectors.toList());
    }

    /**
     * 提取变量中的值，比如 formatParamCode("${1234}"),
     * 会得到结果1234
     *
     * @param paramCode 需要提取的字符串
     * @return String
     * */
    private String formatParamCode(String paramCode){
        if(paramCode == null)
            return "";
        return paramCode.replaceAll("\\$", "")
                .replaceAll("\\{", "")
                .replaceAll("\\}", "");
    }

    /**
     * 使用paramData当中的值替换data当中的变量
     *
     * @param data 需要提取的字符串
     * @param paramData 需要替换的值
     * @return String
     * */
    private String composeMessage(String data, Map<String,String> paramData){
        String regex = "\\$\\{(.+?)\\}";
        Pattern pattern = Pattern.compile(regex);
        Matcher matcher = pattern.matcher(data);
        StringBuffer msg = new StringBuffer();
        while (matcher.find()) {
            String key = matcher.group(1);// 键名
            String value = paramData.get(key);// 键值
            if(value == null) {
                value = "";
            } else {
                value = value.replaceAll("\\$", "\\\\\\$");
            }
            matcher.appendReplacement(msg, value);
        }
        matcher.appendTail(msg);
        return msg.toString();
    }

    // 初始化cellList
    private void initCellList(int sheetNo){
        cellList = new ArrayList<>();
        if(examine() && !initSheet(sheetNo))
            return;
        int rn = sheet.getLastRowNum();
        for(int i = 0;i < rn;i++){
            Row row = sheet.getRow(i);
            if(row != null){
                short cn = row.getLastCellNum();
                for (int j = 0;j < cn;j++){
                    cellList.add(row.getCell(j));
                }
            }
        }
    }

    /**
     * 替换掉所有行区域中的所有 ${} 标记
     * valueList对rows中${}替换的顺序是：
     * 从左至右，从上到下
     *
     * @param rows 行区域
     * @param valueList 替换的值
     * */
    private void replaceMark(List<Row> rows,List<String> valueList){
        if (rows == null || valueList == null)
            return;
        rows.forEach(r -> {
            r.forEach(c -> {
                if(c.getCellTypeEnum() == CellType.STRING && "${}".equals(c.getStringCellValue())){
                    if(valueList == null)
                        return;
                    String value = valueList.stream().findFirst().orElse(null);
                    c.setCellValue(value);
                    if(value != null)
                        valueList.remove(valueList.indexOf(value));
                }
            });
        });
    }

    /**
     * 复制Row到sheet中的另一个Row
     *
     * @param fromSheetNo 复制的行所在的sheet
     * @param fromRow 需要复制的行
     * @param toSheetNo 粘贴的行所在的sheet
     * @param toRow 粘贴的行
     * @param copyValueFlag 是否需要复制值
     */
    private Row copyRow(int fromSheetNo,Row fromRow, int toSheetNo,Row toRow, boolean copyValueFlag) {
        if(fromSheetNo < 0 || fromSheetNo > workbook.getNumberOfSheets()
                || toSheetNo < 0 || toSheetNo > workbook.getNumberOfSheets())
            return null;
        if (fromRow == null || toRow == null)
            return toRow;
        // 设置高度
        toRow.setHeight(fromRow.getHeight());
        // 遍历行中的单元格
        fromRow.forEach(c -> {
            Cell newCell = toRow.createCell(c.getColumnIndex());
            copyCell(c, newCell, copyValueFlag);
        });
        Sheet fromSheet = workbook.getSheetAt(fromSheetNo);
        Sheet toSheet = workbook.getSheetAt(toSheetNo);
        // 遍历行当中的所有的合并区域
        List<CellRangeAddress> crds = fromSheet.getMergedRegions();
        if(crds != null && crds.size() > 0){
            crds.forEach(crd -> {
                // 如果当前合并区域的首行为复制的源行
                if(crd.getFirstRow() == fromRow.getRowNum()) {
                    // 创建对应的合并区域
                    CellRangeAddress newCellRangeAddress = new CellRangeAddress(
                            toRow.getRowNum(),
                            (toRow.getRowNum() + (crd.getLastRow() - crd.getFirstRow())),
                            crd.getFirstColumn(),
                            crd.getLastColumn());
                    // 添加合并区域
                    toSheet.addMergedRegion(newCellRangeAddress);
                }
            });
        }
        return toRow;
    }

    /**
     * 复制Cell到sheet中的另一个Cell
     *
     * @param srcCell 需要复制的单元格
     * @param distCell 粘贴的单元格
     * @param copyValueFlag true则连同cell的内容一起复制
     */
    private void copyCell(Cell srcCell, Cell distCell, boolean copyValueFlag) {
        if (srcCell == null || distCell == null)
            return;
        CellStyle newStyle = workbook.createCellStyle();
        // 获取源单元格的样式
        CellStyle srcStyle = srcCell.getCellStyle();
        // 粘贴样式
        newStyle.cloneStyleFrom(srcStyle);
        // 复制字体
        newStyle.setFont(workbook.getFontAt(srcStyle.getFontIndex()));
        // 复制样式
        distCell.setCellStyle(newStyle);
        // 复制评论
        if(srcCell.getCellComment() != null) {
            distCell.setCellComment(srcCell.getCellComment());
        }
        // 不同数据类型处理
        CellType srcCellType = srcCell.getCellTypeEnum();
        distCell.setCellType(srcCellType);
        if(copyValueFlag) {
            if(srcCellType == CellType.NUMERIC) {
                if(DateUtil.isCellDateFormatted(srcCell)) {
                    distCell.setCellValue(srcCell.getDateCellValue());
                } else {
                    distCell.setCellValue(srcCell.getNumericCellValue());
                }
            } else if(srcCellType == CellType.STRING) {
                distCell.setCellValue(srcCell.getRichStringCellValue());
            } else if(srcCellType == CellType.BLANK) {

            } else if(srcCellType == CellType.BOOLEAN) {
                distCell.setCellValue(srcCell.getBooleanCellValue());
            } else if(srcCellType == CellType.ERROR) {
                distCell.setCellErrorValue(srcCell.getErrorCellValue());
            } else if(srcCellType == CellType.FORMULA) {
                distCell.setCellFormula(srcCell.getCellFormula());
            } else {
            }
        }
    }

    private void exception() throws InvalidFormatException, IOException {
        if(ex != null){
            if(ex instanceof InvalidFormatException)
                throw new InvalidFormatException("错误的文件格式");
            else if(ex instanceof IOException)
                throw new IOException();
            else
                return;
        }
    }

    /**
     * 在插入之前移动并且创建行，避免新插入的行覆盖旧行
     *
     * @param sheetNo 需要操作的Sheet的编号
     * @param startRow 移动的Row区间的起始位置
     * @param moveNum 移动的行数
     * */
    private void shiftAndCreateRows(int sheetNo,int startRow,int moveNum){
        if(!examine() || !initSheet(sheetNo)
                || startRow > sheet.getLastRowNum())
            return;

        List<Sheet> sheetList = new ArrayList<>(Arrays.asList(sheets));
        // 复制当前需要操作的sheet到一个临时的sheet
        Sheet tempSheet = workbook.cloneSheet(sheetList.indexOf(sheet));
        // 获取临时sheet在workbook当中的索引
        int tempSheetNo = workbook.getSheetIndex(tempSheet);
        // 得到临时sheet的第一个row的索引
        int firstRowNum = tempSheet.getFirstRowNum();
        // 得到临时sheet的最后一个row的索引
        int lastRowNum = tempSheet.getLastRowNum();
        int size;
        // 当前操作的sheet整体下移
        sheet.shiftRows(firstRowNum,lastRowNum,
                size = (lastRowNum - firstRowNum + moveNum + 1),true,false);
        // 在腾出的空间上添加新的行，这些新的行不会具有任何的合并单元格，所以可以使用copyRow复制
        // 添加新的行之后删除旧的行
        for(int i= firstRowNum;i < size;i++){
            sheet.createRow(i);
        }
        for(int i= firstRowNum;i < lastRowNum - firstRowNum + 1;i++){
            if(i < startRow)
                copyRow(tempSheetNo,tempSheet.getRow(i),sheetNo,sheet.getRow(i),true);
                // 到达需要插入的索引的位置，需要留出moveNum空间的行
            else
                copyRow(tempSheetNo,tempSheet.getRow(i),sheetNo,sheet.getRow(i + moveNum),true);
            Row row = sheet.getRow(i + size);
            if(row != null)
                sheet.removeRow(row);
        }
        // 删除临时的sheet
        workbook.removeSheetAt(tempSheetNo);
    }

    /**
     * 存储Excel
     *
     * @param path 存储路径
     * @throws IOException
     * @throws InvalidFormatException
     */
    public void save(String path) throws IOException, InvalidFormatException {
        exception();
        if(!examine())
            return;
        try (FileOutputStream fos = new FileOutputStream(path)){
            workbook.write(fos) ;
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 返回Excel的字节数组
     *
     * @return byte[]
     */
    public byte[] getBytes(){
        if(!examine())
            return null;
        try(ByteArrayOutputStream ops = new ByteArrayOutputStream()){
            workbook.write(ops);
            return ops.toByteArray();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }

    @Override
    public boolean equals(Object o){
        if(o == this)
            return true;
        if(!(o instanceof ExcelTemplate))
            return false;
        if(!(examine() ^ ((ExcelTemplate)o).examine()))
            return false;
        return path == ((ExcelTemplate)o).path;
    }

    @Override
    public int hashCode(){
        int hash = Objects.hashCode(path);
        return hash >>> 16 ^ hash;
    }

    @Override
    public String toString(){
        return "ExcelTemplate from " + path + " is " +
                (examine() ? "effective" : "invalid");
    }
}