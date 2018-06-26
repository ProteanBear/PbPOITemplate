package xyz.proteanbear.template.annotation;

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
    //Date pattern
    String dateFormat() default "yyyy-MM-dd HH:mm:ss";
    //Is file path
    boolean isFilePath() default false;
    //the base file path
    String baseFilePath() default "";
}