package com.zhang.excel4j.util;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 正则匹配相关工具
 *
 * author : zhangpan
 * date : 2018/1/29 19:06
 */
public class RegularUtil {

    /**
     * 匹配整数，包含小数点后全是0
     */
    private static final String INTEGER_REG = "^(-?\\d+)(\\.0*)?$";

    /**
     * 匹配科学计数法数字
     */
    private static final String E_NOTATION_REG = "^-?[1-9](\\.\\d+)?(E[+-]?\\d+)$";

    /**
     * 正则提取匹配到的内容
     *
     * @param pattern 匹配目标内容
     * @param reg     正则表达式
     * @param group   提取内容索引
     * @return 提取内容集合
     */
    public static List<String> match(String pattern, String reg, int group) {
        List<String> matchGroups = new ArrayList<>();
        Pattern compile = Pattern.compile(reg);
        Matcher matcher = compile.matcher(pattern);
        if (group > matcher.groupCount() || group < 0)
            return null;
        while (matcher.find()) {
            matchGroups.add(matcher.group(group));
        }
        return matchGroups;
    }

    /**
     * 正则提取匹配到的内容,默认提取索引为0
     *
     * @param pattern 匹配目标内容
     * @param reg     正则表达式
     * @return 提取内容集合
     */
    public static String match(String pattern, String reg) {
        String match = null;
        List<String> matches = match(pattern, reg, 0);
        if (null != matches && matches.size() > 0) {
            match = matches.get(0);
        }
        return match;
    }

    /**
     * 整数去小数点和小数部分
     *
     * @param number 处理前字符串
     * @return 处理后字符串
     */
    public static String convertNumByReg(String number) {
        Pattern compile = Pattern.compile(INTEGER_REG);
        Matcher matcher = compile.matcher(number);
        while (matcher.find()) {
            number = matcher.group(1);
        }
        return number;
    }

    /**
     * 科学计数法数据转换
     *
     * @param bigDecimal 科学计数法
     * @return 数据字符串
     */
    public static String matchDoneBigDecimal(String bigDecimal) {
        // 对科学计数法进行处理
        boolean flag = Pattern.matches(E_NOTATION_REG, bigDecimal);
        if (flag) {
            BigDecimal bd = new BigDecimal(bigDecimal);
            bigDecimal = bd.toPlainString();
        }
        return bigDecimal;
    }
}
