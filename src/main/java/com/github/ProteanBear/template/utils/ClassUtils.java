package com.github.ProteanBear.template.utils;

import java.lang.annotation.Annotation;
import java.lang.reflect.*;
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
     * generate title->getMethod map by annotation
     *
     * @param annotationClass the annotation class
     * @param ofClass         the primary class
     * @return the hash map for the annotation object to the method
     * @throws NoSuchMethodException
     * @throws InvocationTargetException
     * @throws IllegalAccessException
     */
    public static final <T extends Annotation> Map<T,Method> titleMapGetMethodBy(Class<T> annotationClass,Class ofClass)
            throws NoSuchMethodException
    {
        Map<T,Method> result=new LinkedHashMap<>();

        //All fields
        Field[] fields=ofClass.getDeclaredFields();
        Field field=null;
        Method valueMethod=null;
        T annotation=null;
        for(int i=0, length=fields.length;i<length;i++)
        {
            //Get current field
            field=fields[i];
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
     * generate title->setMethod map by annotation
     *
     * @param annotationClass the annotation class
     * @param ofClass         the primary class
     * @return the hash map for valueKey to the method
     * @throws NoSuchMethodException
     * @throws InvocationTargetException
     * @throws IllegalAccessException
     */
    public static final Map<String,Method> titleMapSetMethodBy(Class annotationClass,Class ofClass)
            throws NoSuchMethodException, InvocationTargetException, IllegalAccessException
    {
        Map<String,Method> result=new LinkedHashMap<String,Method>();

        //All fields
        Field[] fields=ofClass.getDeclaredFields();
        Field field=null;
        Method valueMethod=null;
        Object annotation=null;
        for(int i=0, length=fields.length;i<length;i++)
        {
            //Get current field
            field=fields[i];
            //Get annotation
            annotation=field.getAnnotation(annotationClass);

            //Annotation is null
            if(annotation==null) continue;
            //Annotation's value method
            valueMethod=annotationClass.getMethod("value");
            //field's setter method
            result.put((String)valueMethod.invoke(annotation),methodSetterOf(field,ofClass));
        }

        return result;
    }

    /**
     * Get setter method of field
     *
     * @param field
     * @param ofClass
     * @return
     * @throws NoSuchMethodException
     */
    public static final Method methodGetterOf(Field field,Class ofClass) throws NoSuchMethodException
    {
        String name=field.getName();
        StringBuilder getMethodName=new StringBuilder("get");
        getMethodName.append(name.substring(0,1).toUpperCase())
                .append(name.substring(1));

        return ofClass.getMethod(getMethodName.toString());
    }

    /**
     * Get setter method of field
     *
     * @param field
     * @param ofClass
     * @return
     * @throws NoSuchMethodException
     */
    public static final Method methodSetterOf(Field field,Class ofClass) throws NoSuchMethodException
    {
        String name=field.getName();
        StringBuilder setMethodName=new StringBuilder("set");
        setMethodName.append(name.substring(0,1).toUpperCase())
                .append(name.substring(1));

        return ofClass.getMethod(setMethodName.toString(),field.getType());
    }
}