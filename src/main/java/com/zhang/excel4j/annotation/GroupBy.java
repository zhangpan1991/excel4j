package com.zhang.excel4j.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 用来在对象的属性上加入的annotation，通过该annotation说明某个属性所对应的分组
 * RetentionPolicy.RUNTIME 指明其策略是运行时策略，在运行时可以通过反射获取到
 * ElementType.FIELD 指明注解作用目标是属性
 *
 * author : zhangpan
 * date : 2018/1/25 18:24
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface GroupBy {

    /**
     * 包含所有的分组
     */
    String ALL = "all";

    /**
     * 最大排序值
     */
    double MAX = 9999.0;

    /**
     * 声明导出分组
     * @return 分组名称
     */
    String[] value();

    /**
     * 标题排序，升序
     * @return 排序值
     */
    double[] order() default {};
}
