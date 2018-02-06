package com.zhang.excel4j.model;

import org.apache.poi.ss.usermodel.*;

import java.util.HashMap;
import java.util.Map;

/**
 * author : zhangpan
 * date : 2018/2/2 18:39
 */
public class ExcelTemplate {

    /**
     * 当前工作簿
     */
    private Workbook workbook;

    /**
     * 模板sheet表
     */
    private Sheet tempSheet;

    /**
     * 当前sheet
     */
    private Sheet sheet;

    /**
     * 当前行
     */
    private Row row;

    /**
     * 当前列下标
     */
    private int currentCol;

    /**
     * 当前行下标
     */
    private int currentRow;

    /**
     * 序号
     */
    private int serialNumber;

    /**
     * 序号列坐标
     */
    private int serialNumberCol = -1;

    /**
     * 默认样式
     */
    private CellStyle defaultStyle;

    /**
     * 指定行样式
     */
    private Map<Integer, CellStyle> appointLineStyle = new HashMap<>();

    /**
     * 单数行样式
     */
    private CellStyle singleLineStyle;

    /**
     * 双数行样式
     */
    private CellStyle doubleLineStyle;

    /**
     * 数据的初始化列数
     */
    private int initCol;

    /**
     * 数据的初始化行数
     */
    private int initRow;

    /**
     * 最后一行的数据
     */
    private int lastRow;

    /**
     * 默认行高
     */
    private float rowHeight;

    public void createRow() {
        if (this.lastRow > this.currentRow && this.currentRow != this.initRow) {
            this.sheet.shiftRows(this.currentRow, this.lastRow, 1, true, true);
            this.lastRow++;
        }
        this.row = this.sheet.createRow(this.currentRow);
        this.row.setHeightInPoints(this.rowHeight);
        this.currentRow++;
        this.currentCol = this.initCol;
    }

    /**
     * 新加单元格
     *
     * @param value 数据
     */
    public void createCell(Object value) {
        Cell cell = this.row.createCell(currentCol);
        // 设置单元格样式
        setCellStyle(cell);
        cell.setCellValue(value.toString());
    }

    /**
     * 设置单元格样式
     * 优先级：单行样式 > 单数行样式 = 双数行样式 > 默认样式
     *
     * @param cell 单元格
     */
    private void setCellStyle(Cell cell) {
        if (this.appointLineStyle.containsKey(cell.getRowIndex())) {
            cell.setCellStyle(this.appointLineStyle.get(cell.getRowIndex()));
            return;
        }
        if (null != this.singleLineStyle && (cell.getRowIndex() % 2 != 0)) {
            cell.setCellStyle(this.singleLineStyle);
            return;
        }
        if (null != this.doubleLineStyle && (cell.getRowIndex() % 2 == 0)) {
            cell.setCellStyle(this.doubleLineStyle);
            return;
        }
        if (null != this.defaultStyle)
            cell.setCellStyle(this.defaultStyle);
    }

    public Workbook getWorkbook() {
        return workbook;
    }

    public void setWorkbook(Workbook workbook) {
        this.workbook = workbook;
    }

    public Sheet getTempSheet() {
        return tempSheet;
    }

    public void setTempSheet(Sheet tempSheet) {
        this.tempSheet = tempSheet;
    }

    public int getSerialNumber() {
        return serialNumber;
    }

    public Sheet getSheet() {
        return sheet;
    }

    public void setSheet(Sheet sheet) {
        this.sheet = sheet;
    }

    public Row getRow() {
        return row;
    }

    public void setRow(Row row) {
        this.row = row;
    }

    public int getCurrentCol() {
        return currentCol;
    }

    public void setCurrentCol(int currentCol) {
        this.currentCol = currentCol;
    }

    public int getCurrentRow() {
        return currentRow;
    }

    public void setCurrentRow(int currentRow) {
        this.currentRow = currentRow;
    }

    public void setSerialNumber(int serialNumber) {
        this.serialNumber = serialNumber;
    }

    public int getSerialNumberCol() {
        return serialNumberCol;
    }

    public void setSerialNumberCol(int serialNumberCol) {
        this.serialNumberCol = serialNumberCol;
    }

    public CellStyle getDefaultStyle() {
        return defaultStyle;
    }

    public void setDefaultStyle(CellStyle defaultStyle) {
        this.defaultStyle = defaultStyle;
    }

    public Map<Integer, CellStyle> getAppointLineStyle() {
        return appointLineStyle;
    }

    public void setAppointLineStyle(Map<Integer, CellStyle> appointLineStyle) {
        this.appointLineStyle = appointLineStyle;
    }

    public CellStyle getSingleLineStyle() {
        return singleLineStyle;
    }

    public void setSingleLineStyle(CellStyle singleLineStyle) {
        this.singleLineStyle = singleLineStyle;
    }

    public CellStyle getDoubleLineStyle() {
        return doubleLineStyle;
    }

    public void setDoubleLineStyle(CellStyle doubleLineStyle) {
        this.doubleLineStyle = doubleLineStyle;
    }

    public int getInitCol() {
        return initCol;
    }

    public void setInitCol(int initCol) {
        this.initCol = initCol;
    }

    public int getInitRow() {
        return initRow;
    }

    public void setInitRow(int initRow) {
        this.initRow = initRow;
    }

    public int getLastRow() {
        return lastRow;
    }

    public void setLastRow(int lastRow) {
        this.lastRow = lastRow;
    }

    public float getRowHeight() {
        return rowHeight;
    }

    public void setRowHeight(float rowHeight) {
        this.rowHeight = rowHeight;
    }
}
