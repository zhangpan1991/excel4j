package com.zhang.excel4j.handler;

import com.zhang.excel4j.common.WorkbookType;
import com.zhang.excel4j.model.ExcelHeader;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.List;

/**
 * author : zhangpan
 * date : 2018/1/26 11:05
 */
public class ExcelHandler {

    /**
     * 根据注解导出一张工作表的工作簿
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
        sheetWithAnnotation(workbook, data, clazz, sheetName, groupName, isWriteHeader);
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
     * 根据注解创建一张工作表
     *
     * @param workbook      工作簿
     * @param data          数据
     * @param clazz         处理对象
     * @param sheetName     工作表名
     * @param groupName     分组名
     * @param isWriteHeader 是否写入表头
     * @throws Exception 异常
     */
    private static void sheetWithAnnotation(Workbook workbook, List<?> data, Class clazz, String sheetName, String groupName, boolean isWriteHeader) throws Exception {
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
