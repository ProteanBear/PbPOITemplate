package xyz.proteanbear.template.utils;

import java.io.File;

/**
 * File extension name
 *
 * @author ProteanBear
 */
public enum FileSuffix
{
    EXCEL_XLSX(".xlsx"),
    EXCEL_XLS(".xls"),
    WORD_DOC(".doc"),
    WORD_DOCX(".docx");

    /**
     *
     */
    private String suffix;

    /**
     * @param suffix the file suffix
     */
    FileSuffix(String suffix)
    {
        this.suffix=suffix;
    }

    /**
     * @param name file name
     * @return the file suffix
     */
    public boolean check(String name)
    {
        return name!=null && (name.endsWith(suffix));
    }

    /**
     * Get the file type suffix.
     *
     * @param file file object
     * @return file type suffix
     */
    public static FileSuffix getBy(File file)
    {
        String fileName=file.getName();
        if(EXCEL_XLSX.check(fileName)) return EXCEL_XLSX;
        if(EXCEL_XLS.check(fileName)) return EXCEL_XLS;
        if(WORD_DOC.check(fileName)) return WORD_DOC;
        if(WORD_DOCX.check(fileName)) return WORD_DOCX;
        return null;
    }
}