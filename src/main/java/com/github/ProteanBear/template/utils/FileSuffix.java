package com.github.ProteanBear.template.utils;

import java.io.File;

/**
 * File extension name
 *
 * @author ProteanBear
 */
public enum FileSuffix
{
    EXCEL_XLSX(".xlsx"),
    EXCEL_XLS(".xls");

    /**
     *
     */
    private String suffix;

    /**
     * @param suffix
     */
    FileSuffix(String suffix)
    {
        this.suffix=suffix;
    }

    /**
     *
     * @param name
     * @return
     */
    public boolean check(String name)
    {
        return name==null?false:(name.endsWith(suffix));
    }

    /**
     * Get the file type suffix.
     *
     * @param file
     * @return
     */
    public static final FileSuffix getBy(File file)
    {
        String fileName=file.getName();
        if(EXCEL_XLSX.check(fileName)) return EXCEL_XLSX;
        if(EXCEL_XLS.check(fileName)) return EXCEL_XLS;
        return null;
    }
}