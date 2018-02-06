package com.zhang.excel4j.handler;

import com.zhang.excel4j.common.WorkbookType;
import com.zhang.excel4j.converter.Converter;
import com.zhang.excel4j.converter.DefaultConverter;
import com.zhang.excel4j.model.ExcelHeader;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.lang.reflect.Array;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;
import java.util.Map;

/**
 * author : zhangpan
 * date : 2018/1/26 11:05
 */
public class ExcelHandler {

    /**
     * 通过名称读取Sheet到绑定的对象集合，基于注解
     *
     * @param workbook  工作簿
     * @param clazz     处理对象
     * @param titleLine 标题行数
     * @param limitLine 读取行数量
     * @param sheetName Sheet名称
     * @param <T>       绑定的数据类型
     * @return 已绑定数据对象集合
     * @throws Exception 异常
     */
    public static <T> List<T> readSheets(Workbook workbook, Class<T> clazz, int titleLine, int limitLine, String sheetName) throws Exception {
        // 通过Sheet名获取Sheet索引
        int sheetIndex = workbook.getSheetIndex(sheetName);
        if (sheetIndex == -1) {
            return null;
        }
        return readSheets(workbook, clazz, titleLine, limitLine, sheetIndex);
    }

    /**
     * 读取多张Sheet到绑定的对象集合，基于注解
     *
     * @param workbook     工作簿
     * @param clazz        处理对象
     * @param titleLine    标题行数
     * @param limitLine    读取行数量
     * @param sheetIndexes 工作表索引集合
     * @param <T>          绑定的数据类型
     * @return 已绑定数据对象集合
     * @throws Exception 异常
     */
    public static <T> List<T> readSheets(Workbook workbook, Class<T> clazz, int titleLine, int limitLine, int... sheetIndexes) throws Exception {
        // 无Sheet索引，取所有数据
        sheetIndexes = createSheetIndexes(workbook, sheetIndexes);
        if (sheetIndexes.length == 0) {
            return null;
        }
        // 标题列集合
        Map<Integer, ExcelHeader> headerMap = getExcelHeaderMap(workbook, clazz, titleLine, sheetIndexes[0]);
        if (headerMap == null) {
            return null;
        }
        // 单索引读取
        if (sheetIndexes.length == 1) {
            return readSheetBySheetIndex(workbook, clazz, headerMap, titleLine + 1, limitLine, sheetIndexes[0]);
        }
        // 多Sheet索引读取
        List<T> dataList = new ArrayList<>();
        for (int sheetIndex : sheetIndexes ) {
            List<T> list = readSheetBySheetIndex(workbook, clazz, headerMap, titleLine + 1, limitLine, sheetIndex);
            if (list != null) {
                dataList.addAll(list);
            }
        }
        return dataList;
    }

    /**
     * 通过名称读取Sheet数据
     *
     * @param workbook  工作簿
     * @param startLine 开始行数
     * @param limitLine 读取行数量
     * @param sheetName Sheet名称
     * @return 数据集合
     */
    public static List<List<String>> readSheets(Workbook workbook, int startLine, int limitLine, String sheetName) {
        // 通过Sheet名获取Sheet索引
        int sheetIndex = workbook.getSheetIndex(sheetName);
        if (sheetIndex == -1) {
            return null;
        }
        return readSheets(workbook, startLine, limitLine, sheetIndex);
    }

    /**
     * 读取多张Sheet合并数据
     *
     * @param workbook     工作簿
     * @param startLine    开始行数
     * @param limitLine    读取行数量
     * @param sheetIndexes 工作表索引集合
     * @return 数据集合
     */
    public static List<List<String>> readSheets(Workbook workbook, int startLine, int limitLine, int... sheetIndexes) {
        // 无Sheet索引，取所有数据
        sheetIndexes = createSheetIndexes(workbook, sheetIndexes);
        if (sheetIndexes.length == 0) {
            return null;
        }
        // 单索引读取
        if (sheetIndexes.length == 1) {
            return readSheetBySheetIndex(workbook, startLine, limitLine, sheetIndexes[0]);
        }
        // 多Sheet索引读取
        List<List<String>> dataList = new ArrayList<>();
        for (int sheetIndex : sheetIndexes) {
            List<List<String>> list = readSheetBySheetIndex(workbook, startLine, limitLine, sheetIndex);
            if (list != null) {
                dataList.addAll(list);
            }
        }
        return dataList;
    }

    /**
     * Sheet索引为空时，取所有索引
     *
     * @param workbook     工作簿
     * @param sheetIndexes Sheet索引集合
     * @return Sheet索引集合
     */
    private static int[] createSheetIndexes(Workbook workbook, int[] sheetIndexes) {
        if (sheetIndexes.length == 0) {
            // Excel的Sheet数量
            int sheetNum = workbook.getNumberOfSheets();
            if (sheetNum == 0) {
                return sheetIndexes;
            }
            sheetIndexes = new int[sheetNum];
            for (int i = 0; i < sheetNum; i++) {
                sheetIndexes[i] = i;
            }
        }
        return sheetIndexes;
    }

    /**
     * 获取标题列集合
     *
     * @param workbook   工作簿
     * @param clazz      处理对象
     * @param titleLine  标题行数
     * @param sheetIndex Sheet索引
     * @return 标题列集合
     * @throws Exception 异常
     */
    private static Map<Integer, ExcelHeader> getExcelHeaderMap(Workbook workbook, Class<?> clazz, int titleLine, int sheetIndex) throws Exception {
        // 工作表
        Sheet sheet = workbook.getSheetAt(sheetIndex);
        if (sheet == null) {
            return null;
        }
        // 标题行
        Row titleRow = sheet.getRow(titleLine);
        // 标题列对象集合
        Map<Integer, ExcelHeader> headerMap = ColumnHandler.readHeaderMapByTitle(titleRow, clazz);
        if (headerMap == null || headerMap.size() == 0) {
            return null;
        }
        return headerMap;
    }

    /**
     * 读取Sheet到绑定的对象集合，基于注解
     *
     * @param workbook   工作簿
     * @param clazz      处理对象
     * @param headerMap  标题列集合
     * @param startLine  开始行（标题行）数
     * @param limitLine  读取行数量
     * @param sheetIndex 工作表索引
     * @param <T>        数据类型
     * @return 数据集合
     */
    private static <T> List<T> readSheetBySheetIndex(Workbook workbook, Class<T> clazz, Map<Integer, ExcelHeader> headerMap, int startLine, int limitLine, int sheetIndex) throws Exception {
        // 工作表
        Sheet sheet = workbook.getSheetAt(sheetIndex);
        if (sheet == null) {
            return null;
        }
        limitLine = limitLine <= 0 ? Integer.MAX_VALUE : limitLine;
        int totalLine = ((long) startLine + limitLine) > Integer.MAX_VALUE ? Integer.MAX_VALUE : (startLine + limitLine);
        // 结束行
        int endLine = sheet.getLastRowNum() > (totalLine - 1) ? (totalLine - 1) : sheet.getLastRowNum();
        List<T> data = new ArrayList<>();
        for (int i = startLine; i <= endLine; i++) {
            Row row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            T obj = clazz.newInstance();
            boolean empty = true;
            for (Map.Entry<Integer, ExcelHeader> entry : headerMap.entrySet()) {
                // 单元格数据字符串
                String valString = ColumnHandler.getCellValue(row.getCell(entry.getKey()));
                // 判断是否为空
                if (valString != null && !"".equals(valString.trim())) {
                    empty = false;
                }
                ExcelHeader header = entry.getValue();
                Object value;
                // 数据转换
                Converter converter = header.getConverter();
                if (converter != null && converter.getClass() != DefaultConverter.class) {
                    value = converter.execRead(valString);
                } else {
                    // 默认数据转换
                    value = ColumnHandler.str2TargetClass(valString, header.getFiledClazz());
                }
                // 对象赋值
                ColumnHandler.copyProperty(obj, header.getFiled(), value);
            }
            // 判断是否为空行
            if (!empty) {
                data.add(obj);
            }
        }
        return data;
    }

    /**
     * 读取Sheet数据
     *
     * @param workbook   工作簿
     * @param startLine  开始行（标题行）数
     * @param limitLine  读取行数量
     * @param sheetIndex 工作表索引
     * @return 数据集合
     */
    private static List<List<String>> readSheetBySheetIndex(Workbook workbook, int startLine, int limitLine, int sheetIndex) {
        List<List<String>> data = new ArrayList<>();
        // 工作表
        Sheet sheet = workbook.getSheetAt(sheetIndex);
        limitLine = limitLine <= 0 ? Integer.MAX_VALUE : limitLine;
        int totalLine = ((long) startLine + limitLine) > Integer.MAX_VALUE ? Integer.MAX_VALUE : (startLine + limitLine);
        // 结束行
        int endLine = sheet.getLastRowNum() > (totalLine - 1) ? (totalLine - 1) : sheet.getLastRowNum();
        for (int i = startLine; i <= endLine; i++) {
            List<String> rows = new ArrayList<>();
            Row row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            boolean empty = true;
            for (Cell cell : row) {
                String value = ColumnHandler.getCellValue(cell);
                // 判断是否为空
                if (value != null && !"".equals(value.trim())) {
                    empty = false;
                }
                rows.add(value);
            }
            // 判断是否为空行
            if (!empty) {
                data.add(rows);
            }
        }
        return data;
    }

    /**
     * 导出一张工作表，基于注解
     *
     * @param data          数据
     * @param clazz         处理对象
     * @param sheetName     工作表名
     * @param groupName     分组名
     * @param workbookType  工作簿类型
     * @return 工作簿
     * @throws Exception 异常
     */
    public static Workbook exportWorkbook(List<?> data, Class<?> clazz, String sheetName, String groupName, WorkbookType workbookType) throws Exception {
        Workbook workbook = createWorkbook(workbookType);
        createSheet(workbook, data, clazz, sheetName, groupName);
        return workbook;
    }

    /**
     * 导出一张工作表
     *
     * @param data         数据
     * @param header       表头
     * @param sheetName    Sheet名
     * @param workbookType 工作簿类型
     * @return 工作簿
     */
    public static Workbook exportWorkbook(List<?> data, List<String> header, String sheetName, WorkbookType workbookType) {
        Workbook workbook = createWorkbook(workbookType);
        createSheet(workbook, data, header, sheetName);
        return workbook;
    }

    /**
     * 创建工作簿
     *
     * @param workbookType 工作簿类型
     * @return 工作簿
     */
    private static Workbook createWorkbook(WorkbookType workbookType) {
        Workbook workbook;
        switch (workbookType) {
            case XSSF:
                // 07版
                workbook = new XSSFWorkbook();
                break;
            case HSSF:
                // 97版
                workbook = new HSSFWorkbook();
                break;
            default:
                // 默认使用07版
                workbook = new XSSFWorkbook();
        }
        return workbook;
    }

    /**
     * 创建一张工作表，基于注解
     *
     * @param workbook      工作簿
     * @param data          数据
     * @param clazz         处理对象
     * @param sheetName     工作表名
     * @param groupName     分组名
     * @throws Exception 异常
     */
    private static void createSheet(Workbook workbook, List<?> data, Class<?> clazz, String sheetName, String groupName) throws Exception {
        // 创建一张工作表
        Sheet sheet;
        if (sheetName != null && !"".equals(sheetName)) {
            sheet = workbook.createSheet(sheetName);
        } else {
            sheet = workbook.createSheet();
        }
        // 获取表头列表
        List<ExcelHeader> headers;
        if (groupName != null && !"".equals(groupName)) {
            headers = ColumnHandler.getExcelHeaderList(groupName, clazz);
        } else {
            headers = ColumnHandler.getExcelHeaderList(clazz);
        }
        // 创建一行
        Row row = sheet.createRow(0);
        // 写入标题
        for (int i = 0; i < headers.size(); i++) {
            row.createCell(i).setCellValue(headers.get(i).getTitle());
        }
        Object obj;
        for (int i = 0; i < data.size(); i++) {
            // 创建一行（跳过标题行）
            row = sheet.createRow(i + 1);
            obj = data.get(i);
            for (int j = 0; j < headers.size(); j++) {
                if (obj instanceof Map) {
                    // Map数据
                    row.createCell(j).setCellValue(ColumnHandler.getValueByMap((Map) obj, headers.get(j).getFiled(), headers.get(j).getConverter()));
                } else {
                    // 处理对象数据
                    row.createCell(j).setCellValue(ColumnHandler.getValueByAttribute(obj, headers.get(j).getFiled(), headers.get(j).getConverter()));
                }
            }
        }
    }

    /**
     * 创建一张工作表
     *
     * @param workbook  工作簿
     * @param data      数据
     * @param header    表头
     * @param sheetName Sheet名
     */
    private static void createSheet(Workbook workbook, List<?> data, List<String> header, String sheetName) {
        // 创建一张工作表
        Sheet sheet;
        if (sheetName != null && !"".equals(sheetName)) {
            sheet = workbook.createSheet(sheetName);
        } else {
            sheet = workbook.createSheet();
        }
        int rowIndex = 0;
        if (null != header && header.size() > 0) {
            // 写标题
            Row row = sheet.createRow(rowIndex++);
            for (int i = 0; i < header.size(); i++) {
                row.createCell(i, CellType.STRING).setCellValue(header.get(i));
            }
        }
        for (Object object : data) {
            Row row = sheet.createRow(rowIndex++);
            if (object.getClass().isArray()) {
                // 数组
                for (int i = 0; i < Array.getLength(object); i++) {
                    row.createCell(i, CellType.STRING).setCellValue(Array.get(object, i).toString());
                }
            } else if (object instanceof Collection) {
                // Collection集合
                int i = 0;
                for (Object item : (Collection) object) {
                    row.createCell(i++, CellType.STRING).setCellValue(item.toString());
                }
            } else {
                row.createCell(0, CellType.STRING).setCellValue(object.toString());
            }
        }
    }
}
