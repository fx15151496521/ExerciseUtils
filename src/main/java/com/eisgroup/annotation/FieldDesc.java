package com.eisgroup.annotation;

import java.lang.annotation.*;

/**
 * @Description: 记录字段信息
 * @Date: 2019/10/31 18:14
 */
@Documented
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Inherited
public @interface FieldDesc {
    String value() default "";
}
