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
        if(!examine() || sheetNo < 0 || sheetNo > sheets.length - 1)
            return false;
        int sheetNum = workbook.getNumberOfSheets();
        sheets = new Sheet[sheetNum];
        for(int i = 0;i < sheetNum;i++){
            if(i == sheetNo)
                sheet = workbook.getSheetAt(i);
            sheets[i] = workbook.getSheetAt(i);
        }
        sheet = workbook.getSheetAt(sheetNo);
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
                        row = copyRow(sheetNo,sheet.getRow(fromRowStartIndex + i),sheetNo,toRow,true,true);
                    else
                        row = copyRow(sheetNo,sheet.getRow(fromRowStartIndex + i + f),sheetNo,toRow,true,true);
                    temp.add(row);
                }
                // 使用传入的值覆盖${}
                replaceMark(temp,areaValues.get(key));
                rows.addAll(temp);
                n++;
            }
            if(delRowTemp){
                if(toRowIndex >= fromRowEndIndex)
                    removeRowArea(sheetNo,fromRowStartIndex,fromRowEndIndex);
                else
                    removeRowArea(sheetNo,fromRowStartIndex + f,fromRowEndIndex + f);
            }
        }
        return rows.size();
    }

    /**
     * 使用一个已经存在的列区域作为模板，
     * 从sheet的toColumnIndex行开始插入这段列区域,
     * areaValue会从上至下，从左至右的替换掉列区域
     * 中值为 ${} 的单元格的值
     *
     * @param sheetNo 需要操作的Sheet的编号
     * @param fromColumnStartIndex 模板row区域的开始索引
     * @param fromColumnEndIndex 模板row区域的结束索引
     * @param toColumnIndex 开始插入的row索引
     * @param areaValues 替换模板row区域的${}值
     * @param delColumnTemp 是否删除模板row区域
     * @return int 插入的列数量
     * @throws IOException
     * @throws InvalidFormatException
     * */
    public int addColumnByExist(int sheetNo,int fromColumnStartIndex, int fromColumnEndIndex,int toColumnIndex,
                                LinkedHashMap<Integer,LinkedList<String>> areaValues, boolean delColumnTemp)
            throws InvalidFormatException, IOException{
        exception();
        if(!examine()
                || !initSheet(sheetNo)
                || fromColumnStartIndex > fromColumnEndIndex
                || toColumnIndex < 0)
            return 0;
        // 合并区域的列的数量
        int areaNum;
        List<Integer> n = new ArrayList<>();
        n.add(0);
        if(areaValues != null){
            int f = areaValues.size() * (areaNum = (fromColumnEndIndex - fromColumnStartIndex + 1));
            // 创建空白的列
            shiftAndCreateColumns(sheetNo,toColumnIndex-1,f);
            // 获取所有合并区域
            List<CellRangeAddress> crds = sheet.getMergedRegions();
            // 读取需要插入的数据
            for (Integer key:areaValues.keySet()){
                for(int i = 0;i < areaNum;i++){
                    // 获取插入的位置
                    int position = toColumnIndex + n.get(0) * areaNum + i;
                    // 插入的列的位置是在复制区域之后
                    if(toColumnIndex >= fromColumnStartIndex)
                        copyColumn(sheetNo,fromColumnStartIndex + i,sheetNo,position,true);
                        // 插入的列的位置是在复制区域之前
                    else
                        copyColumn(sheetNo,fromColumnStartIndex + i + f,sheetNo,position,true);
                }
                // 复制源列的合并区域到新添加的列
                if(crds != null){
                    crds.forEach(crd -> {
                        // 列偏移量
                        int offset = toColumnIndex - fromColumnStartIndex + areaNum * n.get(0);
                        // 合并区域的宽度
                        int rangeAreaNum = crd.getLastColumn() - crd.getFirstColumn() + 1;
                        // 原合并区域的首列
                        int firstColumn = crd.getFirstColumn();
                        // 需要添加的合并区域首列
                        int addFirstColumn = firstColumn + offset;
                        // 根据插入的列的位置是在复制区域之前还是之后
                        // firstColumn和addFirstColumn分配不同的值
                        firstColumn = toColumnIndex >= fromColumnStartIndex ? firstColumn : firstColumn - f;
                        addFirstColumn = toColumnIndex >= fromColumnStartIndex ? addFirstColumn : toColumnIndex + areaNum * n.get(0);
                        if(firstColumn == fromColumnStartIndex){
                            if(rangeAreaNum > areaNum){
                                mergedRegion(sheetNo,
                                        crd.getFirstRow(),
                                        crd.getLastRow(),
                                        addFirstColumn,
                                        addFirstColumn + areaNum - 1);
                            }
                            else {
                                mergedRegion(sheetNo,
                                        crd.getFirstRow(),
                                        crd.getLastRow(),
                                        addFirstColumn,
                                        addFirstColumn + rangeAreaNum - 1);
                            }
                        }
                    });
                }
                // 填充${}
                List<String> fillValues = areaValues.get(key);
                if (fillValues == null || fillValues.size() == 0)
                    continue;
                List<Cell> needFillCells;
                initCellList(sheetNo);
                needFillCells = cellList;
                // 获取所有的值为${}单元格
                needFillCells = needFillCells.stream().filter(c -> {
                    if(c != null && c.getCellTypeEnum() == CellType.STRING && "${}".equals(c.getStringCellValue()))
                        return true;
                    return false;
                }).collect(Collectors.toList());
                if (needFillCells == null)
                    continue;
                // 所有的${}单元格按照列从小到大，行从小到达的顺序排序
                needFillCells.sort((c1,c2) -> {
                    if (c1 == null && c2 == null) {
                        return 0;
                    }
                    if (c1 == null) {
                        return 1;
                    }
                    if (c2 == null) {
                        return -1;
                    }
                    if(c1.getColumnIndex() > c2.getColumnIndex())
                        return 1;
                    else if(c1.getColumnIndex() < c2.getColumnIndex())
                        return -1;
                    else {
                        if(c1.getRowIndex() > c2.getRowIndex())
                            return 1;
                        else if(c1.getRowIndex() < c2.getRowIndex())
                            return -1;
                        else
                            return 0;
                    }
                });
                needFillCells
                        .stream()
                        .filter(c -> {
                            if(c == null)
                                return false;
                            // 筛选出当前需要填充的单元格
                            return c.getColumnIndex() >= toColumnIndex + areaNum * n.get(0)
                                    && c.getColumnIndex() <= toColumnIndex + areaNum * (n.get(0) + 1);
                        }).forEach(c -> {
                    if(fillValues.size() > 0){
                        // 设置为列的首行，再移除掉首行的值
                        c.setCellValue(fillValues.stream().findFirst().orElse(""));
                        fillValues.remove(0);
                    }
                });
                n.replaceAll(i -> i + 1);
            }
            if(delColumnTemp){
                if(toColumnIndex >= fromColumnStartIndex)
                    removeColumnArea(sheetNo,fromColumnStartIndex,fromColumnEndIndex);
                else
                    removeColumnArea(sheetNo,fromColumnStartIndex + f,fromColumnEndIndex + f);
            }
        }
        return n.get(0);
    }

    /**
     * 使用一个已经存在的列区域作为模板，
     * 从sheet的toColumnIndex行开始插入这段列区域,
     * areaValue会从上至下，从左至右的替换掉列区域
     * 中值为 ${} 的单元格的值
     *
     * @param sheetNo 需要操作的Sheet的编号
     * @param fromColumnStartIndex 模板row区域的开始索引
     * @param fromColumnEndIndex 模板row区域的结束索引
     * @param toColumnIndex 开始插入的row索引
     * @param copyNum 复制数量
     * @param delColumnTemp 是否删除模板row区域
     * @return int 插入的列数量
     * @throws IOException
     * @throws InvalidFormatException
     * */
    public int addColumnByExist(int sheetNo,int fromColumnStartIndex, int fromColumnEndIndex,int toColumnIndex,
                                int copyNum, boolean delColumnTemp)
            throws InvalidFormatException, IOException{
        LinkedHashMap<Integer, LinkedList<String>> map = new LinkedHashMap<>();
        for(int i = 1;i <= copyNum;i++){
            map.put(i,new LinkedList<>());
        }
        return addColumnByExist(sheetNo,fromColumnStartIndex,fromColumnEndIndex,toColumnIndex,map,delColumnTemp);
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
    public int fillVariable(int sheetNo,Map<String,String> fillValues)
            throws IOException, InvalidFormatException {
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
            Cell cell = findCells(sheetNo,s -> {
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
     * 根据行坐标和列坐标定位到单元格，并且使用value填充单元格
     *
     * @param sheetNo 需要操作的Sheet的编号
     * @param rowIndex 填充的值
     * @param columnIndex 填充的值
     * @return boolean 是否成功
     * @throws IOException
     * @throws InvalidFormatException
     **/
    public boolean fillByCoordinate(int sheetNo,int rowIndex,int columnIndex,String value)
            throws IOException, InvalidFormatException {
        exception();
        if(!initSheet(sheetNo))
            return false;
        Row row = sheet.getRow(rowIndex);
        if(row == null)
            return false;
        Cell cell = row.getCell(columnIndex);
        if(cell == null)
            return false;
        if(cell.getCellTypeEnum() == CellType.STRING
                || cell.getCellTypeEnum() == CellType.BLANK)
            cell.setCellValue(value == null ? "" : value);
        return true;
    }

    /**
     * 根据断言predicate查找sheet当中符合条件的cell
     *
     * @param sheetNo 需要操作的Sheet的编号
     * @param predicate 筛选的断言
     * @return List<Cell> 符合条件的Cell
     * */
    public List<Cell> findCells(int sheetNo,Predicate<String> predicate){
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
            if(r != null){
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
            }
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
     * @param needMerged 是否需要合并单元格
     */
    private Row copyRow(int fromSheetNo,Row fromRow, int toSheetNo,Row toRow, boolean copyValueFlag,boolean needMerged) {
        if(fromSheetNo < 0 || fromSheetNo > workbook.getNumberOfSheets()
                || toSheetNo < 0 || toSheetNo > workbook.getNumberOfSheets())
            return null;
        if (fromRow == null)
            return null;
        if(toRow == null){
            Sheet sheet = workbook.getSheetAt(toSheetNo);
            if(sheet == null)
                return null;
            toRow = sheet.createRow(fromRow.getRowNum());
            if(toRow == null)
                return null;
        }
        // 设置高度
        toRow.setHeight(fromRow.getHeight());
        // 遍历行中的单元格
        for(Cell c:fromRow){
            Cell newCell = toRow.createCell(c.getColumnIndex());
            copyCell(c, newCell, copyValueFlag);
        }
        // 如果需要合并
        if(needMerged){
            Sheet fromSheet = workbook.getSheetAt(fromSheetNo);
            Sheet toSheet = workbook.getSheetAt(toSheetNo);
            // 遍历行当中的所有的合并区域
            List<CellRangeAddress> crds = fromSheet.getMergedRegions();
            if(crds != null && crds.size() > 0){
                for(CellRangeAddress crd : crds){
                    // 如果当前合并区域的首行为复制的源行
                    if(crd.getFirstRow() == fromRow.getRowNum()) {
                        // 创建对应的合并区域
                        CellRangeAddress newCellRangeAddress = new CellRangeAddress(
                                toRow.getRowNum(),
                                (toRow.getRowNum() + (crd.getLastRow() - crd.getFirstRow())),
                                crd.getFirstColumn(),
                                crd.getLastColumn());
                        // 添加合并区域
                        safeMergedRegion(toSheetNo,newCellRangeAddress);
                    }
                }
            }
        }
        return toRow;
    }

    /**
     * 复制sheet中列的另一列
     *
     * @param fromSheetNo 复制的行所在的sheet
     * @param fromColumnIndex 需要复制的行索引
     * @param toSheetNo 粘贴的行所在的sheet
     * @param toColumnIndex 粘贴的行
     * @param copyValueFlag 是否需要复制值
     */
    private void copyColumn(int fromSheetNo,int fromColumnIndex,int toSheetNo,
                            int toColumnIndex,boolean copyValueFlag) {
        if(fromSheetNo < 0 || fromSheetNo > workbook.getNumberOfSheets()
                || toSheetNo < 0 || toSheetNo > workbook.getNumberOfSheets())
            return;
        Sheet fromSheet = workbook.getSheetAt(fromSheetNo);
        Sheet toSheet = workbook.getSheetAt(toSheetNo);
        for(int i = 0;i < fromSheet.getLastRowNum();i++){
            Row fromRow = fromSheet.getRow(i);
            Row toRow = toSheet.getRow(i);
            if(fromRow == null)
                continue;
            if(toRow == null)
                toRow = toSheet.createRow(i);
            if(toRow == null)
                continue;
            // 设置高度
            toRow.setHeight(fromRow.getHeight());
            Cell srcCell = fromRow.getCell(fromColumnIndex);
            Cell distCell = toRow.getCell(toColumnIndex);
            if(srcCell == null)
                continue;
            if(distCell == null)
                distCell = toRow.createCell(toColumnIndex);
            // 设置列宽
            toSheet.setColumnWidth(toColumnIndex,fromSheet.getColumnWidth(fromColumnIndex));
            copyCell(srcCell,distCell,copyValueFlag);
        }
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

        // 获取源单元格的样式
        CellStyle srcStyle = srcCell.getCellStyle();
        // 复制样式
        distCell.setCellStyle(srcStyle);

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

    /**
     * 合并单元格区域，本方法是安全的操作，在出现合并冲突的时候，
     * 分割合并区域，然后最大限度的合并冲突区域
     *
     * 使用此方法而不是采用addMergedRegion()和
     * addMergedRegionUnsafe()合并单元格区间，
     * 因为此方法会自行解决合并区间冲突，避免报错或者生成
     * 无法打开的excel
     *
     * @param sheetNo 需要操作的Sheet的编号
     * @param firstRow 开始行
     * @param lastRow 结束行
     * @param firstCol 开始列
     * @param lastCol 结束列
     * */
    public void mergedRegion(int sheetNo,int firstRow, int lastRow, int firstCol, int lastCol){
        if(firstRow > lastRow || firstCol > lastCol)
            return;
        CellRangeAddress address = new CellRangeAddress(firstRow,lastRow,firstCol,lastCol);
        safeMergedRegion(sheetNo,address);
    }

    /**
     * 合并单元格区域，本方法是安全的操作，在出现合并冲突的时候，
     * 分割合并区域，然后最大限度的合并冲突区域
     *
     * @param sheetNo 需要操作的Sheet的编号
     * @param rangeAddress 合并的单元格区域
     * */
    private void safeMergedRegion(int sheetNo,CellRangeAddress rangeAddress){
        if(!examine() || !initSheet(sheetNo) || rangeAddress == null)
            return;
        // 获取所有合并的区域
        List<CellRangeAddress> crds = sheet.getMergedRegions();
        if(crds == null)
            return;
        // 获取描述单元格区域的坐标，
        // 在首行和首列，坐标等于行编号，
        // 在末行和末列，坐标等于行编号加1
        int firstRow = rangeAddress.getFirstRow();
        int lastRow = rangeAddress.getLastRow() + 1;
        int firstColumn = rangeAddress.getFirstColumn();
        int lastColumn = rangeAddress.getLastColumn() + 1;
        // 查找冲突的单元格区域
        CellRangeAddress conflictRange = crds.stream()
                .filter(crd -> {
                    // 获取单元格区域的坐标
                    int cFirstRow = crd.getFirstRow();
                    int cLastRow = crd.getLastRow() + 1;
                    int cFirstColumn = crd.getFirstColumn();
                    int cLastColumn = crd.getLastColumn()  + 1;
                    // 每个合并单元格区域看成一个长方形
                    // 计算两个长方形中心的X坐标的距离
                    float xDistance = (float)(lastColumn + firstColumn)/2
                            - (float)(cLastColumn + cFirstColumn)/2;
                    // 每个合并单元格区域看成一个长方形
                    // 计算两个长方形中心的Y坐标的距离
                    float yDistance = (float)(lastRow + firstRow)/2
                            - (float)(cLastRow + cFirstRow)/2;
                    // 获取距离的绝对值
                    xDistance = xDistance >= 0 ? xDistance : -xDistance;
                    yDistance = yDistance >= 0 ? yDistance : -yDistance;
                    // 如果两个合并区域相交了，返回true
                    if(xDistance < ((float)(lastColumn - firstColumn)/2 + (float)(cLastColumn - cFirstColumn)/2)
                            && yDistance < ((float)(lastRow - firstRow)/2 + (float)(cLastRow - cFirstRow)/2))
                        return true;
                    return false;
                })
                .findFirst()
                .orElse(null);
        // 如果没有查找到冲突的区域，直接合并
        if(conflictRange == null){
            if(examineRange(rangeAddress))
                sheet.addMergedRegion(rangeAddress);
        }
        // 如果合并区域冲突了，分离新增的合并区域
        List<CellRangeAddress> splitRangeAddr = splitRangeAddress(conflictRange,rangeAddress);
        if(splitRangeAddr != null)
            splitRangeAddr.forEach(sra -> safeMergedRegion(sheetNo,sra));
    }

    /**
     * 如果插入的目标合并区域target和sheet中已存在的合并区域source冲突，
     * 把target分割成多个合并区域，这些合并区域都不会和source冲突
     *
     * @param source 已经存在的合并单元格区域
     * @param target 新增的合并单元格区域
     * @return target分离之后的合并单元格列表
     * */
    private List<CellRangeAddress> splitRangeAddress(CellRangeAddress source,CellRangeAddress target){
        List<CellRangeAddress> splitRangeAddr = null;
        if(source == null || target == null)
            return null;
        // 获取source区域的坐标
        int sFirstRow = source.getFirstRow();
        int sLastRow = source.getLastRow() + 1;
        int sFirstColumn = source.getFirstColumn();
        int sLastColumn = source.getLastColumn() + 1;
        // 获取target区域的坐标
        int tFirstRow = target.getFirstRow();
        int tLastRow = target.getLastRow() + 1;
        int tFirstColumn = target.getFirstColumn();
        int tLastColumn = target.getLastColumn() + 1;

        while(true){
            if(splitRangeAddr == null)
                splitRangeAddr = new ArrayList<>();
            // 如果target被切分得无法越过source合并区域，退出循环
            if(tFirstRow >= sFirstRow && tLastRow <= sLastRow
                    && tFirstColumn >= sFirstColumn && tLastColumn <= sLastColumn)
                break;
            // 只考虑Y坐标，当source的最大Y坐标sLastRow在开区间(tFirstRow,tLastRow)
            if(sLastRow > tFirstRow && sLastRow < tLastRow){
                CellRangeAddress address =
                        new CellRangeAddress(sLastRow,tLastRow - 1,tFirstColumn,tLastColumn - 1);
                tLastRow = sLastRow;
                if(examineRange(address))
                    splitRangeAddr.add(address);
            }
            // 只考虑Y坐标，当source的最小Y坐标sFirstRow在开区间(tFirstRow,tLastRow)
            if(sFirstRow > tFirstRow && sFirstRow < tLastRow){
                CellRangeAddress address =
                        new CellRangeAddress(tFirstRow,sFirstRow - 1,tFirstColumn,tLastColumn - 1);
                tFirstRow = sFirstRow;
                if(examineRange(address))
                    splitRangeAddr.add(address);
            }
            // 只考虑X坐标，当source的最小X坐标sFirstColumn在开区间(tFirstColumn,tLastColumn)
            if(sFirstColumn > tFirstColumn && sFirstColumn < tLastColumn){
                CellRangeAddress address =
                        new CellRangeAddress(tFirstRow,tLastRow - 1,tFirstColumn,sFirstColumn - 1);
                tFirstColumn = sFirstColumn;
                if(examineRange(address))
                    splitRangeAddr.add(address);
            }
            // 只考虑X坐标，当source的最大X坐标sLastColumn在开区间(tFirstColumn,tLastColumn)
            if(sLastColumn > tFirstColumn && sLastColumn < tLastColumn){
                CellRangeAddress address =
                        new CellRangeAddress(tFirstRow,tLastRow - 1,sLastColumn,tLastColumn - 1);
                tLastColumn = sLastColumn;
                if(examineRange(address))
                    splitRangeAddr.add(address);
            }
        }
        return splitRangeAddr;
    }

    // 检查合并区域
    private boolean examineRange(CellRangeAddress address){
        if(address == null || !examine())
            return false;
        int firstRowNum = address.getFirstRow();
        int lastRowNum = address.getLastRow();
        int firstColumnNum = address.getFirstColumn();
        int lastColumnNum = address.getLastColumn();
        if(firstRowNum == lastRowNum && firstColumnNum == lastColumnNum)
            return false;
        return true;
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
     * 把sheet[sheetNo]当中所有的行从startRow位置开始，
     * 全部下移moveNum数量的位置，并且在腾出的空间当中创建新行
     *
     * 应该使用本方法而不是采用sheet.shiftRows()和sheet.createRow()，
     * 主要是因为插入一段行的时候会进行如下步骤：
     * 第一：使用shiftRows腾出空间
     * 第二：使用createRow(position)从position开始创建行
     * 但是这样，后面下移的行的合并单元格会部分消失，
     * 并且新创建的行的合并单元格并没有消失，这是因为sheet当中的
     * 大于position的CellRangeAddress并没有跟着下移。
     * 而使用本方法下移并且在中间自动插入行，新插入的行不会含有任何合并单元格，
     * 并且原来的合并单元格也不会消失。
     *
     * @param sheetNo 需要操作的Sheet的编号
     * @param startRow 移动的Row区间的起始位置
     * @param moveNum 移动的行数
     * */
    public void shiftAndCreateRows(int sheetNo,int startRow,int moveNum){
        if(!examine() || !initSheet(sheetNo)
                || startRow > sheet.getLastRowNum())
            return;

        // 复制当前需要操作的sheet到一个临时的sheet
        Sheet tempSheet = workbook.cloneSheet(sheetNo);
        // 获取临时sheet在workbook当中的索引
        int tempSheetNo = workbook.getSheetIndex(tempSheet);
        // 得到临时sheet的第一个row的索引
        int firstRowNum = tempSheet.getFirstRowNum();
        // 得到临时sheet的最后一个row的索引
        int lastRowNum = tempSheet.getLastRowNum();
        if(!clearSheet(sheetNo)){
            return;
        }
        for(int i= firstRowNum;i < lastRowNum - firstRowNum + moveNum + 1;i++){
            sheet.createRow(i);
        }
        for(int i= firstRowNum;i < lastRowNum - firstRowNum + 1;i++){
            if(i < startRow)
                copyRow(tempSheetNo,tempSheet.getRow(i),sheetNo,sheet.getRow(i),true,true);
                // 到达需要插入的索引的位置，需要留出moveNum空间的行
            else
                copyRow(tempSheetNo,tempSheet.getRow(i),sheetNo,sheet.getRow(i + moveNum),true,true);
        }
        settingColumnWidth(tempSheetNo,sheetNo);
        // 删除临时的sheet
        workbook.removeSheetAt(tempSheetNo);
    }

    /**
     * 把sheet[sheetNo]当中所有的列从startColumn位置开始，
     * 全部右移moveNum数量的位置，并且在腾出的空间当中创建新列
     *
     * @param sheetNo 需要操作的Sheet的编号
     * @param startColumn 移动的列区间的起始位置
     * @param moveNum 移动的列数
     * */
    public void shiftAndCreateColumns(int sheetNo,int startColumn,int moveNum){
        if(!examine() || !initSheet(sheetNo))
            return;

        // 复制当前需要操作的sheet到一个临时的sheet
        Sheet tempSheet = workbook.cloneSheet(sheetNo);
        // 获取临时sheet在workbook当中的索引
        int tempSheetNo = workbook.getSheetIndex(tempSheet);
        // 得到临时sheet的第一个row的索引
        int firstRowNum = tempSheet.getFirstRowNum();
        // 得到临时sheet的最后一个row的索引
        int lastRowNum = tempSheet.getLastRowNum();

        if(!clearSheet(sheetNo)){
            return;
        }

        for(int i = firstRowNum;i < lastRowNum;i++){
            Row row = tempSheet.getRow(i);
            if(row != null){
                for(int j = 0;j < moveNum;j++){
                    row.createCell(row.getLastCellNum() + moveNum);
                }
                for(int j = 0;j < row.getLastCellNum();j++){
                    if(j <= startColumn)
                        copyColumn(tempSheetNo,j,sheetNo,j,true);
                    else
                        copyColumn(tempSheetNo,j,sheetNo,j + moveNum,true);
                }
            }
        }
        List<CellRangeAddress> crds = tempSheet.getMergedRegions();
        if(crds == null)
            return;
        crds.forEach(crd -> {
            int firstColumn;
            int lastColumn;
            if((lastColumn = crd.getLastColumn()) <= startColumn)
                safeMergedRegion(sheetNo,crd);
            else if((firstColumn = crd.getFirstColumn()) <= startColumn){
                if(lastColumn > startColumn){
                    CellRangeAddress range = new CellRangeAddress(crd.getFirstRow(),crd.getLastRow(),firstColumn,startColumn);
                    if(examineRange(range))
                        safeMergedRegion(sheetNo,range);
                    range = new CellRangeAddress(crd.getFirstRow(),crd.getLastRow(),
                            startColumn + moveNum + 1,lastColumn + moveNum);
                    if(examineRange(range))
                        safeMergedRegion(sheetNo,range);
                }
            }
            else if(firstColumn > startColumn){
                CellRangeAddress range = new CellRangeAddress(crd.getFirstRow(),crd.getLastRow(),
                        firstColumn + moveNum,lastColumn + moveNum);
                if(examineRange(range))
                    safeMergedRegion(sheetNo,range);
            }
        });
        // 删除临时的sheet
        workbook.removeSheetAt(tempSheetNo);
    }

    /**
     * 移除掉行区域
     *
     * @param sheetNo 需要操作的Sheet的编号
     * @param startRow 起始行
     * @param endRow 结束行
     * */
    public void removeRowArea(int sheetNo,int startRow,int endRow){
        if(!examine() || !initSheet(sheetNo) || startRow > endRow)
            return;

        // 复制当前需要操作的sheet到一个临时的sheet
        Sheet tempSheet = workbook.cloneSheet(sheetNo);
        // 获取临时sheet在workbook当中的索引
        int tempSheetNo = workbook.getSheetIndex(tempSheet);
        // 得到临时sheet的第一个row的索引
        int firstRowNum = tempSheet.getFirstRowNum();
        // 得到临时sheet的最后一个row的索引
        int lastRowNum = tempSheet.getLastRowNum();
        // 清空sheet
        if(!clearSheet(sheetNo)){
            return;
        }

        int delNum = endRow - startRow + 1;
        for(int i = firstRowNum;i < lastRowNum;i++){
            Row fromRow = tempSheet.getRow(i);
            Row toRow =  sheet.createRow(i);
            if(i < startRow)
                copyRow(tempSheetNo,fromRow,sheetNo,toRow,true,false);
            else
                copyRow(tempSheetNo,tempSheet.getRow(i + delNum),sheetNo,toRow,true,false);
        }
        List<CellRangeAddress> crds = tempSheet.getMergedRegions();
        if(crds == null)
            return;
        crds.forEach(crd -> {
            if(crd != null){
                int firstMergedRow = crd.getFirstRow();
                int lastMergedRow = crd.getLastRow();
                int firstMergedColumn = crd.getFirstColumn();
                int lastMergedClolunm = crd.getLastColumn();
                if(lastMergedRow < startRow)
                    safeMergedRegion(sheetNo,crd);
                else if(lastMergedRow >= startRow){
                    if(lastMergedRow <= endRow){
                        if(firstMergedRow < startRow){
                            mergedRegion(sheetNo,firstMergedRow,startRow - 1,firstMergedColumn,lastMergedClolunm);
                        }
                    }
                    else if(lastMergedRow > endRow){
                        if(firstMergedRow < startRow){
                            mergedRegion(sheetNo,firstMergedRow,lastMergedRow - delNum,firstMergedColumn,lastMergedClolunm);
                        }
                        else if(firstMergedRow >= startRow && firstMergedRow <= endRow){
                            mergedRegion(sheetNo,endRow + 1 - delNum,lastMergedRow - delNum,firstMergedColumn,lastMergedClolunm);
                        }
                        else if(firstMergedRow > endRow){
                            mergedRegion(sheetNo,firstMergedRow - delNum,lastMergedRow - delNum,firstMergedColumn,lastMergedClolunm);
                        }
                    }
                }
            }
        });
        settingColumnWidth(tempSheetNo,sheetNo);
        // 删除临时的sheet
        workbook.removeSheetAt(tempSheetNo);
    }

    /**
     * 移除掉列区域
     *
     * @param sheetNo 需要操作的Sheet的编号
     * @param startCol 起始列
     * @param endCol 结束列
     * */
    public void removeColumnArea(int sheetNo,int startCol,int endCol){
        if(!examine() || !initSheet(sheetNo) || startCol > endCol)
            return;

        // 复制当前需要操作的sheet到一个临时的sheet
        Sheet tempSheet = workbook.cloneSheet(sheetNo);
        // 获取临时sheet在workbook当中的索引
        int tempSheetNo = workbook.getSheetIndex(tempSheet);
        // 得到临时sheet的第一个row的索引
        int firstRowNum = tempSheet.getFirstRowNum();
        // 得到临时sheet的最后一个row的索引
        int lastRowNum = tempSheet.getLastRowNum();
        int size = lastRowNum - firstRowNum + 1;

        if(!clearSheet(sheetNo)){
            return;
        }

        for(int i = firstRowNum;i < lastRowNum;i++){
            Row row = tempSheet.getRow(i);
            if(row != null){
                for(int j = 0;j < row.getLastCellNum();j++){
                    // 到达删除区间之前正常复制
                    if(j < startCol)
                        copyColumn(tempSheetNo,j,sheetNo,j,true);
                        // 到达删除区间后，跳过区间长度复制
                    else
                        copyColumn(tempSheetNo,j + endCol - startCol + 1,sheetNo,j,true);
                }
            }
        }
        List<CellRangeAddress> crds = tempSheet.getMergedRegions();
        if(crds == null)
            return;
        crds.forEach(crd -> {
            int delColNum = endCol - startCol + 1;
            int firstMergedRow = crd.getFirstRow();
            int lastMergedRow = crd.getLastRow();
            int firstMergedColumn = crd.getFirstColumn();
            int lastMergedClolunm = crd.getLastColumn();
            if(lastMergedClolunm < startCol)
                safeMergedRegion(sheetNo,crd);
            else if(lastMergedClolunm >= startCol){
                if(lastMergedClolunm <= endCol){
                    if(firstMergedColumn < startCol){
                        mergedRegion(sheetNo,firstMergedRow,lastMergedRow,firstMergedColumn,startCol - 1);
                    }
                }
                else if(lastMergedClolunm > endCol){
                    if(firstMergedColumn < startCol){
                        mergedRegion(sheetNo,firstMergedRow,lastMergedRow,firstMergedColumn,lastMergedClolunm - delColNum);
                    }
                    else if(firstMergedColumn >= startCol && firstMergedColumn <= endCol){
                        mergedRegion(sheetNo,firstMergedRow,lastMergedRow,endCol + 1 - delColNum,lastMergedClolunm - delColNum);
                    }
                    else if(firstMergedColumn > endCol){
                        mergedRegion(sheetNo,firstMergedRow,lastMergedRow,firstMergedColumn - delColNum,lastMergedClolunm -delColNum);
                    }
                }
            }
        });
        // 删除临时的sheet
        workbook.removeSheetAt(tempSheetNo);
    }

    private void settingColumnWidth(int sourceSheetNo,int sheetNo){
        if(sourceSheetNo < 0 || sourceSheetNo > workbook.getNumberOfSheets() ||
                sheetNo < 0 || sheetNo > workbook.getNumberOfSheets())
            return;
        List<Row> rows = new ArrayList<>();
        for(int i = sheet.getFirstRowNum();i < sheet.getLastRowNum();i++){
            Row row = sheet.getRow(i);
            if(row != null)
                rows.add(row);
        }
        Row maxColumnRow = rows.stream().max((r1,r2) -> {
            if (r1 == null && r2 == null) {
                return 0;
            }
            if (r1 == null) {
                return 1;
            }
            if (r2 == null) {
                return -1;
            }
            if (r1.getLastCellNum() == r2.getLastCellNum())
                return 0;
            if (r1.getLastCellNum() > r2.getLastCellNum())
                return 1;
            else
                return -1;
        }).filter(r -> r != null).orElse(null);
        if(maxColumnRow != null){
            int maxColumn = maxColumnRow.getLastCellNum();
            for (int i = 0; i < maxColumn; i++) {
                workbook.getSheetAt(sheetNo).setColumnWidth(i,workbook.getSheetAt(sourceSheetNo).getColumnWidth(i));
            }
        }
    }

    /**
     * 清除掉sheet，清除不是删除，只是会清除所有
     * 的列的值和和合并单元格
     *
     * @param sheetNo 需要操作的Sheet的编号
     * @return boolean true-成功 false-失败
     * */
    public boolean clearSheet(int sheetNo){
        if(!examine())
            return false;
        int sheetNum;
        if(sheetNo < 0 || sheetNo > (sheetNum = workbook.getNumberOfSheets()))
            return false;

        for(int i = 0;i < sheetNum;i++){
            if(i == sheetNo){
                String sheetName = workbook.getSheetName(i);
                workbook.removeSheetAt(i);
                workbook.createSheet(sheetName);
            }
            if(i > sheetNo){
                int offset = i - sheetNo;
                String sheetName = workbook.getSheetName(i-offset);
                Sheet newSheet = workbook.cloneSheet(i-offset);
                workbook.removeSheetAt(i-offset);
                workbook.setSheetName(workbook.getSheetIndex(newSheet),sheetName);
            }
        }
        if(!initSheet(sheetNo))
            return false;
        return true;
    }

    /**
     * 存储Excel
     *
     * @param path 存储路径
     * @throws IOException
     * @throws InvalidFormatException
     */
    public void save(String path) throws
            IOException, InvalidFormatException {
        exception();
        if(!examine())
            return;
        try (FileOutputStream fos = new FileOutputStream(path)){
            workbook.write(fos) ;
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

    /**
     * 返回Workbook
     *
     * @return Workbook
     * @throws IOException
     * @throws InvalidFormatException
     * */
    public Workbook getWorkbook()
            throws IOException, InvalidFormatException {
        exception();
        return workbook;
    }

    /**
     * 返回sheet的行数量
     *
     * @param sheetNo 需要操作的Sheet的编号
     * @return int 行数量
     * */
    public int getSheetRowNum(int sheetNo){
        if(!examine() || !initSheet(sheetNo))
            return 0;
        return sheets[sheetNo].getLastRowNum();
    }

    /**
     * 返回excel的缩放率
     *
     * @param zoom 缩放率
     * */
    public void setZoom(int zoom){
        if(!examine() || !initSheet(workbook.getSheetIndex(sheet)))
            return;
        for (int i = 0; i < sheets.length; i++) {
            sheets[i].setZoom(zoom);
        }
    }

    @Override
    public boolean equals(Object o){
        if(o == this)
            return true;
        if(!(o instanceof ExcelTemplate))
            return false;
        if(examine() ^ ((ExcelTemplate)o).examine())
            return false;
        return Objects.equals(path,((ExcelTemplate)o).path);
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