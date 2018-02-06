package com.zhang.excel4j.common;

/**
 * Excel模板自定义属性
 *
 * author : zhangpan
 * date : 2018/1/29 16:37
 */
public interface TemplateConstant {

    // 数据插入起始坐标点
    String DATA_INDEX = "$data_index";
    // 默认样式
    String DEFAULT_STYLE = "$default_style";
    // 当前标记行样式
    String APPOINT_LINE_STYLE = "$appoint_line_style";
    // 单数行样式
    String SINGLE_LINE_STYLE = "$single_line_style";
    // 双数行样式
    String DOUBLE_LINE_STYLE = "$double_line_style";
    // 序号列坐标点
    String SERIAL_NUMBER = "$serial_number_col";
    // 额外数据的标识符
    String EXTEND_DATA_SIGN = "#";
}
