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
     * @return The title of the sheet,using for exporting a excel file.
     */
    String sheetTitle() default "sheet";

    /**
     * <p>Set excel title line's number(start from 0)</p>
     * <p>Start calculate from not null row</p>
     * <p>If not title row,set -1.</p>
     *
     * @return line's number(start from 0)
     */
    int titleLine() default 0;
}