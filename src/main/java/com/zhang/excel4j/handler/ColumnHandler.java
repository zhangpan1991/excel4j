package com.zhang.excel4j.handler;

import com.zhang.excel4j.annotation.Column;
import com.zhang.excel4j.annotation.GroupBy;
import com.zhang.excel4j.common.GroupType;
import com.zhang.excel4j.model.ExcelHeader;
import org.apache.commons.codec.binary.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;

import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * author : zhangpan
 * date : 2018/1/26 11:05
 */
public class ColumnHandler {

    /**
     * getter或setter枚举
     */
    public enum FieldAccessType {
        GETTER, SETTER
    }

    /**
     * 根据对象注解获取文件表头列表
     *
     * @param clz 类型
     * @return 表头列表
     * @throws IllegalAccessException 异常
     * @throws InstantiationException 异常
     */
    public static List<ExcelHeader> getExcelHeaderList(Class<?> clz) throws IllegalAccessException, InstantiationException {
        return getExcelHeaderList(GroupBy.ALL, clz);
    }

    /**
     * 根据分组和对象注解获取文件表头列表
     *
     * @param group 分组
     * @param clz   类型
     * @return 表头列表
     * @throws IllegalAccessException 异常
     * @throws InstantiationException 异常
     */
    public static List<ExcelHeader> getExcelHeaderList(String group, Class<?> clz) throws IllegalAccessException, InstantiationException {
        List<ExcelHeader> headers = new ArrayList<>();
        List<Field> fields = new ArrayList<>();
        // 获取本类及父类中的属性
        for (Class<?> clazz = clz; clazz != Object.class; clazz = clazz.getSuperclass()) {
            fields.addAll(Arrays.asList(clazz.getDeclaredFields()));
        }
        for (Field field : fields) {
            ExcelHeader header;
            // 是否使用ExportField注解
            if (field.isAnnotationPresent(Column.class)) {
                Column exportField = field.getAnnotation(Column.class);
                header = new ExcelHeader(exportField.value(), exportField.order(), exportField.dataType(), exportField.converter().newInstance(), field.getName(), field.getType());
                if (StringUtils.equals(GroupBy.ALL, group) || GroupType.NON.equals(exportField.groupType())) {
                    headers.add(header);
                    continue;
                }
                // 是否使用GroupBy注解
                if (field.isAnnotationPresent(GroupBy.class)) {
                    GroupBy groupBy = field.getAnnotation(GroupBy.class);
                    String[] groups = groupBy.value();
                    double[] orders = groupBy.order();
                    int index = Arrays.asList(groups).indexOf(group);
                    if ((GroupType.MUST.equals(exportField.groupType()) || GroupType.DEFAULT.equals(exportField.groupType())) && index == -1) {
                        continue;
                    }
                    if (index > -1 && orders.length > index) {
                        header.setOrder(orders[index]);
                    }
                } else if (GroupType.MUST.equals(exportField.groupType())) {
                    continue;
                }
                headers.add(header);
            }
        }
        // 排序
        Collections.sort(headers);
        return headers;
    }

    /**
     * 获取单元格内容
     *
     * @param c 单元格
     * @return 单元格内容
     */
    public static String getCellValue(Cell c) {
        String o;
        switch (c.getCellTypeEnum()) {
            case BLANK:
                o = "";
                break;
            case BOOLEAN:
                o = String.valueOf(c.getBooleanCellValue());
                break;
            case FORMULA:
                o = calculationFormula(c);
                break;
            case NUMERIC:
                o = String.valueOf(c.getNumericCellValue());
                o = matchDoneBigDecimal(o);
                o = converNumByReg(o);
                break;
            case STRING:
                o = c.getStringCellValue();
                break;
            default:
                o = null;
                break;
        }
        return o;
    }

    /**
     * 科学计数法数据转换
     *
     * @param bigDecimal 科学计数法
     * @return 数据字符串
     */
    private static String matchDoneBigDecimal(String bigDecimal) {
        // 对科学计数法进行处理
        boolean flg = Pattern.matches("^-?\\d+(\\.\\d+)?(E-?\\d+)?$", bigDecimal);
        if (flg) {
            BigDecimal bd = new BigDecimal(bigDecimal);
            bigDecimal = bd.toPlainString();
        }
        return bigDecimal;
    }

    /**
     * 计算公式结果
     *
     * @param cell 单元格类型为公式的单元格
     * @return 返回单元格计算后的值 格式化成String
     */
    public static String calculationFormula(Cell cell) {
        CellValue cellValue = cell.getSheet().getWorkbook().getCreationHelper()
                .createFormulaEvaluator().evaluate(cell);
        return cellValue.formatAsString();
    }

    /**
     * 通过正则表达式获取有效的数字字符串
     *
     * @param number 字符串
     * @return 数字字符串
     */
    public static String converNumByReg(String number) {
        Pattern compile = Pattern.compile("^(\\d+)(\\.0*)?$");
        Matcher matcher = compile.matcher(number);
        while (matcher.find()) {
            number = matcher.group(1);
        }
        return number;
    }
}
