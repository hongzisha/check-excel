package com.example.bonc.utils;

import com.example.bonc.enumUtils.EnumConstant;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author TanHao
 * @date 2022/10/19 0019
 */
@Slf4j
public class OfficeUtils {

    /**
     * 读取到路径下的excel文件
     * @param fileName  文件名
     * @param flag  标识是周报或者考勤数据
     * @return excel数据
     */
    public static List<Map<Object, Map<Integer, Object>>> creatWorkBook(String fileName, int flag) {

        log.info("判断文件是否存在...");
        File excelFile = new File(fileName);
        if (!excelFile.exists()) {
            log.warn("指定的Excel文件不存在！");
            return null;
        }

        log.info("创建 WorkBook ...");
        String fileType = fileName.substring(fileName.lastIndexOf(".") + 1);
        Workbook workbook;
        try {
            FileInputStream fileInputStream = new FileInputStream(excelFile);
            if (EnumConstant.FORMAT1.getCode().equalsIgnoreCase(fileType)) {
                //生成.xls的excel
                workbook = new HSSFWorkbook(fileInputStream);
            } else if (EnumConstant.FORMAT2.getCode().equalsIgnoreCase(fileType)) {
                //生成.xlsx的excel
                workbook = new XSSFWorkbook(fileInputStream);
            } else {
                log.warn("文件格式不对");
                return null;
            }
            log.info("开始解析 WorkBook...");
            return parseExcelNew(workbook,flag);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }

    /**
     * 获取excel文件的数据
     * @param workbook  excel对象
     * @param flag  标识是周报或者考勤数据
     * @return excel的数据
     */
    private static List<Map<Object, Map<Integer, Object>>> parseExcelNew(Workbook workbook, int flag) {
        List<Map<Object, Map<Integer, Object>>> list = new ArrayList<>();
        for (int sheetNum = 0; sheetNum < workbook.getNumberOfSheets(); sheetNum++) {
            Map<Object, Map<Integer, Object>> contents = new HashMap<>(8);
            // 获取表格
            Sheet sheet = workbook.getSheetAt(sheetNum);
            // 获取第一行，一般是标题
            Row firstRow = sheet.getRow(sheet.getFirstRowNum());
            if (null == firstRow) {
                log.warn("解析Excel失败，在第一行没有读取到任何数据！");
            }
            // 解析每一行的数据，构造数据对象
            //标题下面的数据,数据起始行
            int rowStart = (firstRow != null ? firstRow.getRowNum() : 0) + 1;
            //获取有记录的行数，即：最后有数据的行是第n行，前面有m行是空行没数据，则返回n-m；
            int rowEnd = sheet.getPhysicalNumberOfRows();
            // 获取excel文件中的列数
            int cellNum = firstRow != null ? firstRow.getPhysicalNumberOfCells() : 0;
            if (flag == 1) {
                 contents = function1(rowStart, rowEnd, cellNum, sheet);
            }
            else if (flag == 2){
                 contents = function2(rowStart, rowEnd, cellNum, sheet);
            }
            list.add(contents);
        }
        return list;
    }

    /**
     * 获取考勤数据
     * @param rowStart  开始行
     * @param rowEnd    结束行
     * @param cellNum   列数
     * @param sheet sheet
     * @return 考勤数据
     */
    private static Map<Object, Map<Integer, Object>> function1 (int rowStart, int rowEnd, int cellNum, Sheet sheet) {
        Map<Integer, Object> cellValue = new HashMap<>(8);
        String str = null;
        Map<Object, Map<Integer, Object>> contents = new HashMap<>(8);
        for (int rowNum = rowStart; rowNum < rowEnd; rowNum++) {
            Row row = sheet.getRow(rowNum);
            int flag = 0;
            if (row != null) {
                //处理Cell
                for (int i = 1; i < cellNum + 1; i++) {
                    Cell cell = row.getCell(i - 1);
                    if (cell != null) {
                        String content = convertCellValueToString(cell);
                        // 记录第一行第一列的员工姓名
                        if (rowNum == 1 && i ==1){
                            str = content;
                        }
                        // 当编历到下一个姓名时，保存这个姓名的时间
                        if (null != content && !(str != null && str.equals(content)) && i == 1){
                            contents.put(str, cellValue);
                            str = content;
                            cellValue = new HashMap<>(8);
                        }
                        // 遍历每一列时，将姓名保存
                        if (i == 3){
                            cellValue.put(rowNum, content);
                        }
                        if (StringUtils.isNotBlank(String.valueOf(cell))){
                            flag = 1;
                        }

                        // TODO 处理数据
                    }
                }
            }
            // 当到最后时，保存最后一周的姓名
            if (flag == 1){
                contents.put(str, cellValue);
            }
        }
        return contents;
    }

    /**
     * 获取周报数据
     * @param rowStart  开始行
     * @param rowEnd    结束行
     * @param cellNum   列数
     * @param sheet sheet
     * @return  周报数据
     */
    private static Map<Object, Map<Integer, Object>> function2 (int rowStart, int rowEnd, int cellNum, Sheet sheet) {
        Map<Integer, Object> cellValue = new HashMap<>(8);
        String str = null;
        Map<Object, Map<Integer, Object>> contents = new HashMap<>(8);
        for (int rowNum = rowStart; rowNum < rowEnd; rowNum++) {
            Row row = sheet.getRow(rowNum);
            if (row != null) {
                //处理Cell
                for (int i = 1; i < cellNum + 1; i++) {
                    Cell cell = row.getCell(i - 1);
                    if (cell != null ) {
                        String content = convertCellValueToString(cell);
                        // 记录第一行第一列的时间内容
                        if (rowNum == 1 && i == 1){
                            str = content;
                        }
                        // 当遍历到下一个周的时间时，保存一周的姓名
                        if (StringUtils.isNotBlank(String.valueOf(cell)) && i == 1 && rowNum !=1){
                            contents.put(str,cellValue);
                            str=content;
                            cellValue = new HashMap<>(8);
                        }
                        // 遍历每一列时，将姓名保存
                        if (i == 2){
                            cellValue.put(rowNum, content);
                        }
                        // 当到最后时，保存最后一周的姓名
                        if (rowNum + 1 == rowEnd){
                            contents.put(str, cellValue);
                        }
                        // TODO 处理数据
                    }
                }
            }
        }
        return contents;
    }

    /**
     * 转换excel数据的格式为字符类型
     * @param   cell
     * NUMERIC-数字或者时间,STRING-字符串,BOOLEAN-布尔,BLANK-空值,FORMULA-公式,ERROR-故障
     * @return  返回字符串
     */
    private static String convertCellValueToString(Cell cell) {
        if (cell == null) {
            return null;
        }
        String content = null;
        try {
            switch (cell.getCellTypeEnum()) {
                case NUMERIC:
                    Double doubleValue = cell.getNumericCellValue();
                    // 格式化科学计数法，取一位整数
                    DecimalFormat df = new DecimalFormat("0");
                    content = df.format(doubleValue);
                    break;
                case STRING:
                    content = cell.getStringCellValue();
                    break;
                case BOOLEAN:
                    Boolean booleanValue = cell.getBooleanCellValue();
                    content = String.valueOf(booleanValue);
                    break;
                case FORMULA:
                    content = cell.getCellFormula();
                    break;
                default:
                    break;
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return content;
    }

}
