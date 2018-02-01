package com.zhang.excel4j.handler;

import com.zhang.excel4j.common.WorkbookType;
import com.zhang.excel4j.converter.Converter;
import com.zhang.excel4j.converter.DefaultConverter;
import com.zhang.excel4j.model.ExcelHeader;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * author : zhangpan
 * date : 2018/1/26 11:05
 */
public class ExcelHandler {

    public static void aa(InputStream is) throws Exception {
        Workbook workbook = WorkbookFactory.create(is);
        readSheetWithAnnotationBySheetIndex(workbook, ExcelHeader.class, 0, 0, 0);
    }

    public <T> List<T> readWorkbookWithAnnotation(InputStream is, Class<T> clazz, int startLine, int limitLine, int... sheetIndexes) throws Exception {
        Workbook workbook = WorkbookFactory.create(is);
        List<T> dataList = new ArrayList<>();

        for (int i = 0; i < sheetIndexes.length; i++) {
            List<T> list = readSheetWithAnnotationBySheetIndex(workbook, clazz, startLine, limitLine, sheetIndexes[i]);
            if (list != null) {
                dataList.addAll(list);
            }
        }
        return dataList;
    }

    /**
     * 导出一张工作表的工作簿，基于注解
     *
     * @param data          数据
     * @param clazz         处理对象
     * @param sheetName     工作表名
     * @param groupName     分组名
     * @param isWriteHeader 是否写入表头
     * @param workbookType  工作簿类型
     * @return 工作簿
     * @throws Exception 异常
     */
    public static Workbook exportWorkbookWithAnnotation(List<?> data, Class clazz, String sheetName, String groupName, boolean isWriteHeader, WorkbookType workbookType) throws Exception {
        Workbook workbook = createWorkbook(workbookType);
        createSheetWithAnnotation(workbook, data, clazz, sheetName, groupName, isWriteHeader);
        return workbook;
    }

    /**
     * 读取工作表中的数据，基于注解
     *
     * @param workbook   工作簿
     * @param clazz      处理对象
     * @param startLine  开始行（标题行）数
     * @param limitLine  读取行数量
     * @param sheetIndex 工作表索引
     * @param <T>        数据类型
     * @return 数据集合
     */
    private static <T> List<T> readSheetWithAnnotationBySheetIndex(Workbook workbook, Class<T> clazz,
                                                                   int startLine, int limitLine, int sheetIndex) throws Exception {
        // 工作表
        Sheet sheet = workbook.getSheetAt(sheetIndex);
        // 结束行
        int endLine = sheet.getLastRowNum() > (startLine + limitLine) ? (startLine + limitLine) : sheet.getLastRowNum();
        // 标题行
        Row titleRow = sheet.getRow(startLine);
        // 标题列对象集合
        Map<Integer, ExcelHeader> headerMap = ColumnHandler.readHeaderMapByTitle(titleRow, clazz);
        if (headerMap == null || headerMap.size() == 0) {
            return null;
        }
        List<T> data = new ArrayList<>();
        for (int i = startLine + 1; i <= endLine; i++) {
            Row row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            T obj = clazz.newInstance();
            for (Map.Entry<Integer, ExcelHeader> entry : headerMap.entrySet()) {
                // 单元格数据字符串
                String valString = ColumnHandler.getCellValue(row.getCell(entry.getKey()));
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
            data.add(obj);
        }
        return data;
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
     * @param isWriteHeader 是否写入表头
     * @throws Exception 异常
     */
    private static void createSheetWithAnnotation(Workbook workbook, List<?> data, Class clazz, String sheetName, String groupName, boolean isWriteHeader) throws Exception {
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
        if (isWriteHeader) {
            // 写入标题
            for (int i = 0; i < headers.size(); i++) {
                row.createCell(i).setCellValue(headers.get(i).getTitle());
            }
        }
        Object obj;
        for (int i = 0; i < data.size(); i++) {
            // 创建一行（跳过标题行）
            row = sheet.createRow(i + 1);
            // TODO 对象或Map类型处理
            obj = data.get(i);
            for (int j = 0; j < headers.size(); j++) {
                row.createCell(j).setCellValue(ColumnHandler.getValueByAttribute(obj, headers.get(j).getFiled(), headers.get(j).getConverter()));
            }
        }
    }
}
