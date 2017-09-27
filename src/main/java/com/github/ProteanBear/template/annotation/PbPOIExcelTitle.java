package com.github.ProteanBear.template.annotation;

import java.lang.annotation.*;

/**
 * Custom annotation for mapping title to field
 *
 * @author ProteanBear
 */
@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface PbPOIExcelTitle
{
    //title
    String value();
}