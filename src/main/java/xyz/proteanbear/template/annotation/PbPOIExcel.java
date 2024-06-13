package xyz.proteanbear.template.annotation;

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
    /**
     * The title of the sheet,using for exporting a excel file.
     */
    String sheetTitle() default "sheet";

    /**
     * Set excel title line's number(start from 0)<br/>
     * Start calculate from not null row<br/>
     * If not title row,set -1.
     */
    int titleLine() default 0;
}