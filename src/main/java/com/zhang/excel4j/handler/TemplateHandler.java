package com.zhang.excel4j.handler;

import com.zhang.excel4j.common.TemplateConstant;
import com.zhang.excel4j.model.ExcelHeader;
import com.zhang.excel4j.model.ExcelTemplate;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.IOException;
import java.io.InputStream;
import java.util.List;
import java.util.Map;

/**
 * author : zhangpan
 * date : 2018/1/29 18:04
 */
public class TemplateHandler {

    /**
     * 导出数据到Excel模板中，基于注解
     *
     * @param is            模板输入流
     * @param sheetIndex    模板sheet索引
     * @param data          表数据
     * @param extendData    额外数据
     * @param clazz         处理对象
     * @param groupName     分组名称
     * @param isWriteHeader 是否插入标题行
     * @param sheetName     sheet名称
     * @param isCopySheet   是否复制sheet
     * @return Excel模板
     * @throws Exception 异常
     */
    public static ExcelTemplate exportExcelTemplate(InputStream is, int sheetIndex, List<?> data, Map<String, Object> extendData, Class<?> clazz,
                                                    String groupName, boolean isWriteHeader, String sheetName, boolean isCopySheet) throws Exception {
        sheetIndex = sheetIndex < 0 ? 0 : sheetIndex;
        ExcelTemplate template = loadTemplate(is, sheetIndex, extendData, sheetName, isCopySheet);
        // 获取表头列表
        List<ExcelHeader> headers;
        if (groupName != null && !"".equals(groupName)) {
            headers = ColumnHandler.getExcelHeaderList(groupName, clazz);
        } else {
            headers = ColumnHandler.getExcelHeaderList(clazz);
        }
        if (isWriteHeader) {
            template.createRow();
            template.createSerialNumber(true);
            for (ExcelHeader header : headers) {
                template.createCell(header.getTitle());
            }
        }
        for (Object object : data) {
            template.createRow();
            template.createSerialNumber(false);
            for (ExcelHeader header : headers) {
                if (object instanceof Map) {
                    // Map数据
                    template.createCell(ColumnHandler.getValueByMap((Map) object, header.getFiled(), header.getConverter()));
                } else {
                    // 处理对象数据
                    template.createCell(ColumnHandler.getValueByAttribute(object, header.getFiled(), header.getConverter()));
                }
            }
        }
        return  template;
    }

    /**
     * 导出数据到Excel模板中
     *
     * @param is          模板输入流
     * @param sheetIndex  模板sheet索引
     * @param data        表数据
     * @param extendData  额外数据
     * @param sheetName   sheet名称
     * @param isCopySheet 是否复制sheet
     * @return Excel模板
     * @throws Exception 异常
     */
    public static ExcelTemplate exportExcelTemplate(InputStream is, int sheetIndex, List<?> data, Map<String, Object> extendData, String sheetName, boolean isCopySheet) throws Exception {
        sheetIndex = sheetIndex < 0 ? 0 : sheetIndex;
        ExcelTemplate template = loadTemplate(is, sheetIndex, extendData, sheetName, isCopySheet);
        for (Object object : data) {
            template.createRow();
            template.createSerialNumber(false);
            for (Map.Entry<Integer, String> entry : template.getHeaderMap().entrySet()) {
                if (object instanceof Map) {
                    // Map数据
                    template.createCell(ColumnHandler.getValueByMap((Map) object, entry.getValue(), null), entry.getKey());
                } else {
                    // 处理对象数据
                    template.createCell(ColumnHandler.getValueByAttribute(object, entry.getValue(), null), entry.getKey());
                }
            }
        }
        return template;
    }

    /**
     * 加载Excel模板
     *
     * @param is          模板输入流
     * @param sheetIndex  模板sheet索引
     * @param extendData  额外数据
     * @param sheetName   sheet名称
     * @param isCopySheet 是否复制sheet
     * @return Excel模板
     * @throws IOException            异常
     * @throws InvalidFormatException 异常
     */
    private static ExcelTemplate loadTemplate(InputStream is, int sheetIndex, Map<String, Object> extendData, String sheetName, boolean isCopySheet) throws IOException, InvalidFormatException {
        Workbook workbook = WorkbookFactory.create(is);
        ExcelTemplate excelTemplate = new ExcelTemplate();
        excelTemplate.setWorkbook(workbook);
        // sheet模板
        Sheet tempSheet = workbook.getSheetAt(sheetIndex);
        excelTemplate.setTempSheet(tempSheet);
        excelTemplate.setLastRow(tempSheet.getLastRowNum());
        // 设置序号
        excelTemplate.setSerialNumber(1);

        for (Row row : tempSheet) {
            for (Cell cell : row) {
                if (cell.getCellTypeEnum().equals(CellType.STRING)) {
                    String value = cell.getStringCellValue().trim();
                    // 替换额外数据
                    if (value.startsWith(TemplateConstant.EXTEND_DATA_SIGN)) {
                        if (extendData.containsKey(value.substring(1))) {
                            cell.setCellValue(extendData.get(value.substring(1)).toString());
                        }
                    }
                    // 自定义数据映射表头
                    if (value.startsWith(TemplateConstant.DATA_HEADER_SIGN)) {
                        excelTemplate.getHeaderMap().put(cell.getColumnIndex(), value.substring(1));
                    }
                    // 获取模板数据和样式
                    switch (value) {
                        // 序号列
                        case TemplateConstant.SERIAL_NUMBER:
                            excelTemplate.setSerialNumberCol(cell.getColumnIndex());
                            clearCell(cell);
                            break;
                        // 数据列
                        case TemplateConstant.DATA_INDEX:
                            excelTemplate.setInitCol(cell.getColumnIndex());
                            excelTemplate.setInitRow(cell.getRowIndex());
                            excelTemplate.setCurrentCol(cell.getColumnIndex());
                            excelTemplate.setCurrentRow(cell.getRowIndex());
                            excelTemplate.setRowHeight(row.getHeightInPoints());
                            clearCell(cell);
                            break;
                        // 默认行样式
                        case TemplateConstant.DEFAULT_STYLE:
                            excelTemplate.setDefaultStyle(cell.getCellStyle());
                            clearCell(cell);
                            break;
                        // 单行样式
                        case TemplateConstant.APPOINT_LINE_STYLE:
                            excelTemplate.getAppointLineStyle().put(cell.getRowIndex(), cell.getCellStyle());
                            clearCell(cell);
                            break;
                        // 单数行样式
                        case TemplateConstant.SINGLE_LINE_STYLE:
                            excelTemplate.setSingleLineStyle(cell.getCellStyle());
                            clearCell(cell);
                            break;
                        // 双数行样式
                        case TemplateConstant.DOUBLE_LINE_STYLE:
                            excelTemplate.setDoubleLineStyle(cell.getCellStyle());
                            clearCell(cell);
                            break;
                    }
                }
            }
        }

        Sheet sheet;
        // 写入数据的sheet
        if (isCopySheet) {
            // 复制模板
            sheet = excelTemplate.getWorkbook().createSheet();
            copySheet(tempSheet, sheet, 0, tempSheet.getLastRowNum());
        } else {
            sheet = excelTemplate.getTempSheet();
        }
        // 修改sheet名称
        if (sheetName != null && !"".equals(sheetName)) {
            int index = excelTemplate.getWorkbook().getSheetIndex(sheet);
            excelTemplate.getWorkbook().setSheetName(index, sheetName);
        }
        excelTemplate.setSheet(sheet);
        return excelTemplate;
    }

    /**
     * 清除单元格
     *
     * @param cell 单元格
     */
    private static void clearCell(Cell cell) {
        cell.setCellStyle(null);
        cell.setCellValue("");
    }

    /**
     * 复制sheet
     *
     * @param fromSheet 源sheet
     * @param toSheet   目标sheet
     * @param firstRow  开始行
     * @param lastRow   结束行
     */
    private static void copySheet(Sheet fromSheet, Sheet toSheet, int firstRow, int lastRow) {
        if ((firstRow == -1) || (lastRow == -1) || lastRow < firstRow) {
            return;
        }
        // 复制合并的单元格
        CellRangeAddress region;
        for (int i = 0; i < fromSheet.getNumMergedRegions(); i++) {
            region = fromSheet.getMergedRegion(i);
            if ((region.getFirstRow() >= firstRow) && (region.getLastRow() <= lastRow)) {
                toSheet.addMergedRegion(region);
            }
        }
        Row fromRow;
        Row newRow;
        Cell fromCell;
        Cell newCell;
        // 设置列宽
        for (int i = firstRow; i <= lastRow; i++) {
            fromRow = fromSheet.getRow(i);
            if (fromRow != null) {
                for (int j = fromRow.getLastCellNum(); j >= fromRow.getFirstCellNum(); j--) {
                    toSheet.setColumnWidth(j, fromSheet.getColumnWidth(j));
                    toSheet.setColumnHidden(j, false);
                }
                break;
            }
        }
        // 复制行并填充数据
        for (int i = firstRow; i <= lastRow; i++) {
            fromRow = fromSheet.getRow(i);
            if (fromRow == null) {
                continue;
            }
            newRow = toSheet.createRow(i - firstRow);
            newRow.setHeight(fromRow.getHeight());
            for (int j = fromRow.getFirstCellNum(); j < fromRow.getPhysicalNumberOfCells(); j++) {
                fromCell = fromRow.getCell(j);
                if (fromCell == null) {
                    continue;
                }
                newCell = newRow.createCell(j);
                newCell.setCellStyle(fromCell.getCellStyle());
                switch (fromCell.getCellTypeEnum()) {
                    case BOOLEAN:
                        newCell.setCellValue(fromCell.getBooleanCellValue());
                        break;
                    case FORMULA:
                        newCell.setCellValue(fromCell.getCellFormula());
                        break;
                    case NUMERIC:
                        newCell.setCellValue(fromCell.getNumericCellValue());
                        break;
                    case ERROR:
                        newCell.setCellValue(fromCell.getErrorCellValue());
                        break;
                    default:
                        newCell.setCellValue(fromCell.getRichStringCellValue());
                        break;
                }
            }
        }
    }
}
