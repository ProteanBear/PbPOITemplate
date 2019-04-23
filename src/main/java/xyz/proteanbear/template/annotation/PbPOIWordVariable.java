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
public @interface PbPOIWordVariable
{
    //title
    String value();
}