package com.github.ProteanBear.template.annotation;

import java.lang.annotation.*;

/**
 * Custom annotations mark Excel file attributes.
 *
 * @author ProteanBear
 */
@Target({ElementType.TYPE})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface PbPOIExcel
{
    //The title of the sheet
    String sheetTitle() default "";
}