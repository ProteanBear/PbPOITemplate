package xyz.proteanbear.template.annotation;

import org.apache.poi.common.usermodel.PictureType;

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

    //Is it an image path
    boolean isImagePath() default false;

    PictureType imageType() default PictureType.JPEG;

    String imageDescription() default "";

    int imageWidth() default 480;

    int imageHeight() default 360;
}