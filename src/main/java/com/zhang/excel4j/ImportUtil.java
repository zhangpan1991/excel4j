package com.zhang.excel4j;

import com.zhang.excel4j.handler.ExcelHandler;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;
import java.io.InputStream;
import java.util.List;

/**
 * author : zhangpan
 * date : 2018/1/25 19:21
 */
public class ImportUtil {

    /**
     * 单例模式
     */
    private static ImportUtil importUtil;

    private ImportUtil() {
    }

    public static ImportUtil getInstance() {
        if (importUtil == null) {
            synchronized (ImportUtil.class) {
                if (importUtil == null) {
                    importUtil = new ImportUtil();
                }
            }
        }
        return importUtil;
    }

    public <T> List<T> readExcelWithAnnotation(InputStream is, Class<T> clazz) throws Exception {
        return readExcelWithAnnotation(is, clazz, 0, Integer.MAX_VALUE);
    }

    public <T> List<T> readExcelWithAnnotation(String excelPath, Class<T> clazz) throws Exception {
        return readExcelWithAnnotation(excelPath, clazz, 0, Integer.MAX_VALUE);
    }

    public <T> List<T> readExcelWithAnnotation(InputStream is, Class<T> clazz, int sheetIndex) throws Exception {
        return readExcelWithAnnotation(is, clazz, 0, Integer.MAX_VALUE, sheetIndex);
    }

    public <T> List<T> readExcelWithAnnotation(String excelPath, Class<T> clazz, int sheetIndex) throws Exception {
        return readExcelWithAnnotation(excelPath, clazz, 0, Integer.MAX_VALUE, sheetIndex);
    }

    public <T> List<T> readExcelWithAnnotation(InputStream is, Class<T> clazz, int titleLine, String sheetName) throws Exception {
        return readExcelWithAnnotation(is, clazz, titleLine, (Integer.MAX_VALUE - titleLine), sheetName);
    }

    public <T> List<T> readExcelWithAnnotation(String excelPath, Class<T> clazz, int titleLine, String sheetName) throws Exception {
        return readExcelWithAnnotation(excelPath, clazz, titleLine, (Integer.MAX_VALUE - titleLine), sheetName);
    }

    public <T> List<T> readExcelWithAnnotation(InputStream is, Class<T> clazz, int titleLine, int limitLine, String sheetName) throws Exception {
        Workbook workbook = WorkbookFactory.create(is);
        return ExcelHandler.readSheets(workbook, clazz, titleLine, limitLine, sheetName);
    }

    public <T> List<T> readExcelWithAnnotation(String excelPath, Class<T> clazz, int titleLine, int limitLine, String sheetName) throws Exception {
        Workbook workbook = WorkbookFactory.create(new File(excelPath));
        return ExcelHandler.readSheets(workbook, clazz, titleLine, limitLine, sheetName);
    }

    public <T> List<T> readExcelWithAnnotation(InputStream is, Class<T> clazz, int titleLine, int limitLine, int... sheetIndexes) throws Exception {
        Workbook workbook = WorkbookFactory.create(is);
        return ExcelHandler.readSheets(workbook, clazz, titleLine, limitLine, sheetIndexes);
    }

    public <T> List<T> readExcelWithAnnotation(String excelPath, Class<T> clazz, int titleLine, int limitLine, int... sheetIndexes) throws Exception {
        Workbook workbook = WorkbookFactory.create(new File(excelPath));
        return ExcelHandler.readSheets(workbook, clazz, titleLine, limitLine, sheetIndexes);
    }

    public List<List<String>> readExcel2List(InputStream is, int startLine, int limitLine, String sheetName) throws Exception {
        Workbook workbook = WorkbookFactory.create(is);
        return ExcelHandler.readSheets(workbook, startLine, limitLine, sheetName);
    }

    public List<List<String>> readExcel2List(String excelPath, int startLine, int limitLine, String sheetName) throws Exception {
        Workbook workbook = WorkbookFactory.create(new File(excelPath));
        return ExcelHandler.readSheets(workbook, startLine, limitLine, sheetName);
    }

    public List<List<String>> readExcel2List(InputStream is, int startLine, int limitLine, int... sheetIndexes) throws Exception {
        Workbook workbook = WorkbookFactory.create(is);
        return ExcelHandler.readSheets(workbook, startLine, limitLine, sheetIndexes);
    }

    public List<List<String>> readExcel2List(String excelPath, int startLine, int limitLine, int... sheetIndexes) throws Exception {
        Workbook workbook = WorkbookFactory.create(new File(excelPath));
        return ExcelHandler.readSheets(workbook, startLine, limitLine, sheetIndexes);
    }
}
