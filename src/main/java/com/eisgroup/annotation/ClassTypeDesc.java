package com.eisgroup.annotation;

import java.lang.annotation.*;

/**
 * @Description: 记录实体类型
 * @Date: 2019/10/31 18:14
 */
@Documented
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
@Inherited
public @interface ClassTypeDesc {
    String value() default "";
}
