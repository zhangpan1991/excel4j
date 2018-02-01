package com.zhang.excel4j.annotation;

import com.zhang.excel4j.common.GroupType;
import com.zhang.excel4j.converter.Converter;
import com.zhang.excel4j.converter.DefaultConverter;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 用来在对象的属性上加入的annotation，通过该annotation说明某个属性所对应的标题
 * RetentionPolicy.RUNTIME 指明其策略是运行时策略，在运行时可以通过反射获取到
 * ElementType.FIELD 指明注解作用目标是属性
 *
 * author : zhangpan
 * date : 2018/1/25 16:49
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface Column {

    /**
     * 属性对应的标题名称
     * @return 表头名称
     */
    String value() default "";

    /**
     * 分组类型
     * @return 分组类型
     */
    GroupType groupType() default GroupType.DEFAULT;

    /**
     * 按全局排序值升序排序
     * @return 全局排序值
     */
    double order() default GroupBy.MAX;

    /**
     * 数据转换器
     * @see Converter
     * @return 文件数据转换器
     */
    Class<? extends Converter> converter()
            default DefaultConverter.class;
}
