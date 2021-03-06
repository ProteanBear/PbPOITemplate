package xyz.proteanbear.template.utils;

import xyz.proteanbear.template.annotation.PbPOIExcelTitle;

import java.lang.annotation.Annotation;
import java.lang.reflect.*;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Map;

/**
 * Class related tools.
 *
 * @author ProteanBear
 */
public class ClassUtils
{
    /**
     * generate title — getMethod map by annotation
     *
     * @param annotationClass the annotation class
     * @param ofClass         the primary class
     * @return the hash map for the annotation object to the method
     * @throws NoSuchMethodException No such method
     */
    public static <T extends Annotation> Map<T,Method> titleMapGetMethodBy(Class<T> annotationClass,Class ofClass)
            throws NoSuchMethodException
    {
        Map<T,Method> result=new LinkedHashMap<>();

        //All fields
        Field[] fields=ofClass.getDeclaredFields();
        T annotation;
        for(Field field : fields)
        {
            //Get annotation
            annotation=field.getAnnotation(annotationClass);

            //Annotation is null
            if(annotation==null) continue;
            //field's setter method
            result.put(annotation,methodGetterOf(field,ofClass));
        }

        return result;
    }

    /**
     * generate title — setMethod map by annotation
     *
     * @param annotationClass    the annotation class
     * @param ofClass            the primary class
     * @param titleMapAnnotation the map for title — annotation object
     * @return the hash map for valueKey to the method
     * @throws NoSuchMethodException     No such method
     * @throws InvocationTargetException Invocation target
     * @throws IllegalAccessException    Illegal access
     */
    public static <T extends Annotation> Map<String,Method> titleMapSetMethodBy(Class<T> annotationClass,Class ofClass,
                                                                                Map<String,Object> titleMapAnnotation)
            throws NoSuchMethodException, InvocationTargetException, IllegalAccessException
    {
        Map<String,Method> result=new LinkedHashMap<>();

        //All fields
        Field[] fields=ofClass.getDeclaredFields();
        Method valueMethod;
        Object annotation;
        String method;
        for(Field field : fields)
        {
            //Get annotation
            annotation=field.getAnnotation(annotationClass);

            //Annotation is null
            if(annotation==null) continue;
            //Annotation's value method
            valueMethod=annotationClass.getMethod("value");
            //method
            method=(String)valueMethod.invoke(annotation);
            //field's setter method
            result.put(method,methodSetterOf(field,ofClass));
            //annotation
            titleMapAnnotation.put(method,annotation);
        }

        return result;
    }

    /**
     * generate title — getMethod map by annotation
     *
     * @param annotationClass the annotation class
     * @param data            data objects with template variables declared by annotations
     * @return the hash map for the annotation value() to the data content
     * @throws NoSuchMethodException     No such method
     * @throws InvocationTargetException Invocation target
     * @throws IllegalAccessException    Illegal access
     */
    public static <T extends Annotation> Map<String,Object> dataMapBy(Class<T> annotationClass,Object data)
            throws NoSuchMethodException, InvocationTargetException, IllegalAccessException
    {
        Map<String,Object> result=new HashMap<>();

        //All fields
        Class dataClass=data.getClass();
        Field[] fields=dataClass.getDeclaredFields();
        T annotation;
        Method valueMethod, getMethod;
        String method;
        for(Field field : fields)
        {
            //Get annotation
            annotation=field.getAnnotation(annotationClass);

            //Annotation is null
            if(annotation==null) continue;
            //Annotation's value method
            valueMethod=annotationClass.getMethod("value");
            //method
            method=(String)valueMethod.invoke(annotation);
            //field's getter method
            getMethod=methodGetterOf(field,dataClass);

            result.put(method,getMethod.invoke(data));
        }

        return result;
    }

    /**
     * Get setter method of field
     *
     * @param fieldName The class field name string
     * @param ofClass   The class
     * @return The field's getter method
     * @throws NoSuchMethodException No such method
     * @throws NoSuchFieldException  No such field
     */
    public static Method methodGetterOf(String fieldName,Class ofClass)
            throws NoSuchMethodException, NoSuchFieldException
    {
        return methodGetterOf(ofClass.getDeclaredField(fieldName),ofClass);
    }

    /**
     * Get setter method of field
     *
     * @param field   The class field
     * @param ofClass The class
     * @return The field's getter method
     * @throws NoSuchMethodException No such method
     */
    private static Method methodGetterOf(Field field,Class ofClass) throws NoSuchMethodException
    {
        String name=field.getName();
        String getMethodName="get"+name.substring(0,1).toUpperCase()+
                name.substring(1);
        return ofClass.getMethod(getMethodName);
    }

    /**
     * Get setter method of field
     *
     * @param field   The class field
     * @param ofClass The class
     * @return The field's getter method
     * @throws NoSuchMethodException No such method
     */
    private static Method methodSetterOf(Field field,Class ofClass) throws NoSuchMethodException
    {
        String name=field.getName();
        String setMethodName="set"+name.substring(0,1).toUpperCase()+
                name.substring(1);
        return ofClass.getMethod(setMethodName,field.getType());
    }
}