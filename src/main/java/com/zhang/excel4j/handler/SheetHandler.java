package com.zhang.excel4j.handler;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 * author : zhangpan
 * date : 2018/4/4 17:03
 */
public class SheetHandler {

    /**
     * 清除单元格
     *
     * @param cell 单元格
     */
    public static void clearCell(Cell cell) {
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
    public static void copySheet(Sheet fromSheet, Sheet toSheet, int firstRow, int lastRow) {
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
                newCell.getCellStyle().cloneStyleFrom(fromCell.getCellStyle());
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
