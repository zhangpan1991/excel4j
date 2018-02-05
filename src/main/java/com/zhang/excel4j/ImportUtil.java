package com.zhang.excel4j;

import com.zhang.excel4j.handler.ExcelHandler;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

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
    private volatile static ImportUtil importUtil;

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

    /**
     * 读取excel数据绑定到对象，基于注解
     *
     * @param is    输入流
     * @param clazz 处理对象
     * @param <T>   绑定的数据类型
     * @return 已绑定数据对象集合
     * @throws Exception 异常
     */
    public <T> List<T> readExcel2List(InputStream is, Class<T> clazz) throws Exception {
        return readExcel2List(is, clazz, 0, Integer.MAX_VALUE);
    }

    /**
     * 读取excel数据绑定到对象，基于注解
     *
     * @param is         输入流
     * @param clazz      处理对象
     * @param sheetIndex sheet索引
     * @param <T>        绑定的数据类型
     * @return 已绑定数据对象集合
     * @throws Exception 异常
     */
    public <T> List<T> readExcel2List(InputStream is, Class<T> clazz, int sheetIndex) throws Exception {
        return readExcel2List(is, clazz, 0, Integer.MAX_VALUE, sheetIndex);
    }

    /**
     * 读取excel数据绑定到对象，基于注解
     *
     * @param is        输入流
     * @param clazz     处理对象
     * @param sheetName sheet名称
     * @param <T>       绑定的数据类型
     * @return 已绑定数据对象集合
     * @throws Exception 异常
     */
    public <T> List<T> readExcel2List(InputStream is, Class<T> clazz, String sheetName) throws Exception {
        return readExcel2List(is, clazz, 0, Integer.MAX_VALUE, sheetName);
    }

    /**
     * 读取excel数据绑定到对象，基于注解
     *
     * @param is        输入流
     * @param clazz     处理对象
     * @param titleLine 标题行数
     * @param limitLine 读取行数量
     * @param sheetName sheet名称
     * @param <T>       绑定的数据类型
     * @return 已绑定数据对象集合
     * @throws Exception 异常
     */
    public <T> List<T> readExcel2List(InputStream is, Class<T> clazz, int titleLine, int limitLine, String sheetName) throws Exception {
        Workbook workbook = WorkbookFactory.create(is);
        return ExcelHandler.readSheets(workbook, clazz, titleLine, limitLine, sheetName);
    }

    /**
     * 读取excel数据绑定到对象，基于注解
     *
     * @param is           输入流
     * @param clazz        处理对象
     * @param titleLine    标题行数
     * @param limitLine    读取行数量
     * @param sheetIndexes sheet索引集合
     * @param <T>          绑定的数据类型
     * @return 已绑定数据对象集合
     * @throws Exception 异常
     */
    public <T> List<T> readExcel2List(InputStream is, Class<T> clazz, int titleLine, int limitLine, int... sheetIndexes) throws Exception {
        Workbook workbook = WorkbookFactory.create(is);
        return ExcelHandler.readSheets(workbook, clazz, titleLine, limitLine, sheetIndexes);
    }

    /**
     * 读取excel数据
     *
     * @param is 输入流
     * @return 数据List
     * @throws Exception 异常
     */
    public List<List<String>> readExcel2List(InputStream is) throws Exception {
        return readExcel2List(is, 1, Integer.MAX_VALUE);
    }

    /**
     * 读取excel数据
     *
     * @param is         输入流
     * @param sheetIndex sheet索引
     * @return 数据List
     * @throws Exception 异常
     */
    public List<List<String>> readExcel2List(InputStream is, int sheetIndex) throws Exception {
        return readExcel2List(is, 1, Integer.MAX_VALUE, sheetIndex);
    }

    /**
     * 读取excel数据
     *
     * @param is        输入流
     * @param sheetName sheet名称
     * @return 数据List
     * @throws Exception 异常
     */
    public List<List<String>> readExcel2List(InputStream is, String sheetName) throws Exception {
        return readExcel2List(is, 1, Integer.MAX_VALUE, sheetName);
    }

    /**
     * 读取excel数据
     *
     * @param is        输入流
     * @param startLine 开始行数
     * @param limitLine 读取行数量
     * @param sheetName sheet名称
     * @return 数据List
     * @throws Exception 异常
     */
    public List<List<String>> readExcel2List(InputStream is, int startLine, int limitLine, String sheetName) throws Exception {
        Workbook workbook = WorkbookFactory.create(is);
        return ExcelHandler.readSheets(workbook, startLine, limitLine, sheetName);
    }

    /**
     * 读取excel数据
     *
     * @param is           输入流
     * @param startLine    开始行数
     * @param limitLine    读取行数量
     * @param sheetIndexes sheet索引集合
     * @return 数据List
     * @throws Exception 异常
     */
    public List<List<String>> readExcel2List(InputStream is, int startLine, int limitLine, int... sheetIndexes) throws Exception {
        Workbook workbook = WorkbookFactory.create(is);
        return ExcelHandler.readSheets(workbook, startLine, limitLine, sheetIndexes);
    }
}
