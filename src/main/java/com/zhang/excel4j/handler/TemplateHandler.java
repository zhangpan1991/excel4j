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

    public static ExcelTemplate exportExcelTemplate(InputStream is, int sheetIndex, List<?> data, Map<String, Object> extendData, Class<?> clazz,
                                                    String groupName, boolean isWriteHeader, String sheetName, boolean isCopySheet)
            throws IOException, InvalidFormatException, InstantiationException, IllegalAccessException {
        sheetIndex = sheetIndex < 0 ? 0 : sheetIndex;
        ExcelTemplate template = loadTemplate(is, sheetIndex, extendData);
        Sheet sheet;
        // 写入数据的sheet
        if (isCopySheet) {
            sheet = template.getWorkbook().createSheet();
        } else {
            sheet = template.getTempSheet();
        }
        // 修改sheet名称
        if (sheetName != null && !"".equals(sheetName)) {
            int index = template.getWorkbook().getSheetIndex(sheet);
            template.getWorkbook().setSheetName(index, sheetName);
        }
        template.setSheet(sheet);
        // 获取表头列表
        List<ExcelHeader> headers;
        if (groupName != null && !"".equals(groupName)) {
            headers = ColumnHandler.getExcelHeaderList(groupName, clazz);
        } else {
            headers = ColumnHandler.getExcelHeaderList(clazz);
        }
        if (isWriteHeader) {
            template.createRow();
            for (ExcelHeader header : headers) {
                // TODO 插入数据
            }
        }
        return  template;
    }

    private static ExcelTemplate loadTemplate(InputStream is, int sheetIndex, Map<String, Object> extendData) throws IOException, InvalidFormatException {
        Workbook workbook = WorkbookFactory.create(is);
        ExcelTemplate excelTemplate = new ExcelTemplate();
        excelTemplate.setWorkbook(workbook);
        // sheet模板
        Sheet tempSheet = workbook.getSheetAt(sheetIndex);
        excelTemplate.setTempSheet(tempSheet);
        excelTemplate.setLastRow(tempSheet.getLastRowNum());

        for (Row row : tempSheet) {
            for (Cell cell : row) {
                if (cell.getCellTypeEnum().equals(CellType.STRING)) {
                    String value = cell.getStringCellValue().trim();
                    // 替换额外数据
                    if (value.startsWith(TemplateConstant.EXTEND_DATA_SIGN)) {
                        cell.setCellValue(extendData.get(value.substring(1)).toString());
                    }
                    // 获取模板数据和样式
                    value = value.toLowerCase();
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
        return excelTemplate;
    }

    private static void clearCell(Cell cell) {
        cell.setCellStyle(null);
        cell.setCellValue("");
    }

    public static void copySheet(Sheet fromsheet, Sheet tosheet, int firstrow, int lasttrow) {
        if ((firstrow == -1) || (lasttrow == -1) || lasttrow < firstrow) {
            return;
        }
        // 复制合并的单元格
        CellRangeAddress region;
        for (int i = 0; i < fromsheet.getNumMergedRegions(); i++) {
            region = fromsheet.getMergedRegion(i);
            if ((region.getFirstRow() >= firstrow) && (region.getLastRow() <= lasttrow)) {
                tosheet.addMergedRegion(region);
            }
        }
        Row fromRow;
        Row newRow;
        Cell fromCell;
        Cell newCell;
        // 设置列宽
        for (int i = firstrow; i <= lasttrow; i++) {
            fromRow = fromsheet.getRow(i);
            if (fromRow != null) {
                for (int j = fromRow.getLastCellNum(); j >= fromRow.getFirstCellNum(); j--) {
                    tosheet.setColumnWidth(j, fromsheet.getColumnWidth(j));
                    tosheet.setColumnHidden(j, false);
                }
                break;
            }
        }
        // 复制行并填充数据
        for (int i = firstrow; i <= lasttrow; i++) {
            fromRow = fromsheet.getRow(i);
            if (fromRow == null) {
                continue;
            }
            newRow = tosheet.createRow(i - firstrow);
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
