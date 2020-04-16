package xyz.proteanbear.template;

import org.apache.poi.ss.usermodel.*;
import xyz.proteanbear.template.annotation.PbPOIExcel;
import xyz.proteanbear.template.annotation.PbPOIExcelTitle;
import xyz.proteanbear.template.exception.FileSuffixNotSupportException;
import xyz.proteanbear.template.utils.ClassUtils;
import xyz.proteanbear.template.utils.FileSuffix;
import xyz.proteanbear.template.utils.Hex26Utils;
import xyz.proteanbear.template.utils.StringUtils;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * The tools is created for easy use of Apache POI.
 * This tool is used for reading and writing in Excel.
 *
 * @author ProteanBear
 */
public class PbPOIExcelTemplate
{
    /**
     * Set excel title line's number(start from 0)<br/>
     * Start calculate from not null row<br/>
     * If not title row,set -1.
     */
    private int titleLine = 0;

    /**
     * Read excel file and write to list of T class
     *
     * @param excelFile   The file .xls or .xlsx
     * @param returnClass the class type when return
     * @return the object list of `returnClass`
     * @throws IOException               io exception
     * @throws NoSuchMethodException     no such method
     * @throws IllegalAccessException    illegal access exception
     * @throws InvocationTargetException invocation target exception
     * @throws InstantiationException    instantiation error
     * @throws ParseException            parse error
     */
    public List<?> readFrom(File excelFile, Class<?> returnClass)
            throws IOException, NoSuchMethodException, IllegalAccessException,
            InvocationTargetException, InstantiationException, ParseException
    {
        //Load Excel File
        Workbook workbook = WorkbookFactory.create(excelFile);

        //title->Annotation map
        Map<String, Object> titleAnnotationMap = new HashMap<>();
        //Get returnClass title->setMethod map by annotation
        Map<String, Method> titleMethodMap = ClassUtils
                .titleMapSetMethodBy(PbPOIExcelTitle.class, returnClass, titleAnnotationMap);

        //Record index->title map
        Map<String, String> indexTitleMap = new HashMap<>();

        //All sheets
        Sheet sheet;
        Row row;
        List<Object> result = new ArrayList<>();
        int pageNum = workbook.getNumberOfSheets();
        for (int page = 0; page < pageNum; page++)
        {
            //Get current sheet
            sheet = workbook.getSheetAt(page);

            //All rows be not null
            Iterator<Row> rowItr = sheet.rowIterator();
            int rowNum = 0;
            while (rowItr.hasNext())
            {
                //Get current row
                row = rowItr.next();
                int colNum = row.getLastCellNum();

                //Set index title for not title excel
                if (titleLine == -1 && rowNum == 0)
                {
                    //Set title to "A……Z" mode
                    for (int i = 0; i < colNum; i++)
                    {
                        indexTitleMap.put(i + "", Hex26Utils.from(i + 1));
                    }
                }
                //If current row is title line
                //,record into index->title map
                else if (titleLine == rowNum)
                {
                    for (int index = 0; index < colNum; index++)
                    {
                        indexTitleMap.put(index + "", valueOf(row.getCell(index)) + "");
                    }
                    rowNum++;
                    continue;
                }

                //Create new tClass Object instance
                Object object = returnClass.getDeclaredConstructor()
                                           .newInstance();
                //All cells include null cell
                for (int index = 0; index < colNum; index++)
                {
                    //Set field content,index->title->setMethod
                    Method method = titleMethodMap.get(indexTitleMap.get(index + ""));
                    PbPOIExcelTitle annotation =
                            (PbPOIExcelTitle) titleAnnotationMap.get(indexTitleMap.get(index + ""));
                    if (method == null) continue;
                    Object value = transform(
                            valueOf(row.getCell(index)),
                            method.getParameterTypes()[0],
                            annotation
                    );
                    if (value == null) continue;

                    method.invoke(object, value);
                }
                //Insert into list result.
                result.add(object);

                rowNum++;
            }
        }

        return result;
    }

    /**
     * Writes the specified data list to an Excel file.
     *
     * @param excelFile the excel file
     * @param data      the specified data list.Multiple sets of data to generate multiple sheets.
     * @throws IOException                   io exception
     * @throws FileSuffixNotSupportException file suffix is not supported
     */
    public void writeTo(File excelFile, List<?>... data) throws IOException, FileSuffixNotSupportException
    {
        //File exists, delete the old file
        if (excelFile.exists()) excelFile.delete();
        excelFile.createNewFile();

        writeTo(
                FileSuffix.getBy(excelFile),
                new FileOutputStream(excelFile),
                data
        );
    }

    /**
     * Writes the specified data list to an Excel file.
     *
     * @param fileSuffix   File's suffix
     * @param outputStream Output stream
     * @param data         the specified data list.Multiple sets of data to generate multiple sheets.
     * @throws IOException io exception
     */
    public void writeTo(FileSuffix fileSuffix, OutputStream outputStream, List<?>... data) throws IOException
    {
        //Declare a workbook
        Workbook workbook;
        switch (fileSuffix)
        {
            case EXCEL_XLSX:
                workbook = WorkbookFactory.create(true);
                break;
            case EXCEL_XLS:
            default:
                workbook = WorkbookFactory.create(false);
        }

        //Create the sheets
        PbPOIExcel pbPOIExcelAnnotation;
        Map<PbPOIExcelTitle, Method> getMethodMap;
        for (List<?> oneDataList : data)
        {
            if (oneDataList.isEmpty()) continue;

            //Get the class corresponding annotation
            Class curClass = oneDataList.get(0)
                                        .getClass();
            pbPOIExcelAnnotation = oneDataList.get(0)
                                              .getClass()
                                              .getAnnotation(PbPOIExcel.class);
            if (pbPOIExcelAnnotation == null) continue;

            //Create a sheet
            String sheetTitle = pbPOIExcelAnnotation.sheetTitle();
            Sheet sheet = StringUtils.isBlank(sheetTitle) ? workbook.createSheet() : workbook.createSheet(sheetTitle);
            int curRow = 0, index = 0;

            //Generate the table title line
            try
            {
                getMethodMap = ClassUtils.titleMapGetMethodBy(PbPOIExcelTitle.class, curClass);

                Row row = sheet.createRow(curRow++);
                for (PbPOIExcelTitle pbPOIExcelTitle : getMethodMap.keySet())
                {
                    Cell cell = row.createCell(index);
                    set(cell, pbPOIExcelTitle.value(), pbPOIExcelTitle);
                    index++;
                }
            }
            catch (Exception e)
            {
                e.printStackTrace();
                continue;
            }

            //Generate the table content line
            for (Object anOneDataList : oneDataList)
            {
                Row row = sheet.createRow(curRow++);
                index = 0;

                for (PbPOIExcelTitle pbPOIExcelTitle : getMethodMap.keySet())
                {
                    Cell cell = row.createCell(index);
                    try
                    {
                        set(
                                cell,
                                getMethodMap.get(pbPOIExcelTitle)
                                            .invoke(anOneDataList),
                                pbPOIExcelTitle
                        );
                    }
                    catch (Exception e)
                    {
                        e.printStackTrace();
                        set(cell, "", pbPOIExcelTitle);
                    }
                    index++;
                }
            }
        }

        workbook.write(outputStream);
    }

    /**
     * return cell's value object
     *
     * @param cell Excel
     * @return Object, eg. String/Double/Date/Boolean/null
     */
    public Object valueOf(Cell cell)
    {
        Object result = null;
        if (cell == null) return null;

        //Get cell's type
        CellType type = cell.getCellType();
        //handle different type
        switch (type)
        {
            //String
            case STRING:
            {
                result = cell.getRichStringCellValue()
                             .toString()
                             .trim();
                break;
            }
            //Number(Double)
            case NUMERIC:
                //Formula(Double)
            case FORMULA:
            {
                //Date check
                //Return java.util.Date object
                if (org.apache.poi.ss.usermodel.DateUtil.isCellDateFormatted(cell))
                {
                    result = cell.getDateCellValue();
                }
                //Normal Number
                //Return Double object
                else
                {
                    result = cell.getNumericCellValue();
                }
                break;
            }
            //Boolean
            case BOOLEAN:
            {
                result = cell.getBooleanCellValue();
                break;
            }
            //其他
            default:
        }

        return result;
    }

    /**
     * Transform value to target class,and handle exceptions
     *
     * @param value      value object
     * @param toClass    target class
     * @param annotation annotation
     * @param <T>        target class
     * @return value transformed
     */
    @SuppressWarnings("unchecked")
    private <T> T transform(Object value, Class<T> toClass, PbPOIExcelTitle annotation) throws ParseException
    {
        //value is null
        if (value == null) return null;
        //Date -> String
        if (toClass.isAssignableFrom(String.class))
        {
            return (T) ((value instanceof Date)
                        ? (new SimpleDateFormat(annotation.dateFormat()).format(value))
                        : (String.valueOf(value)));
        }
        //String -> Date
        if (toClass.isAssignableFrom(Date.class))
        {
            return (T) ((value instanceof String)
                        ? (new SimpleDateFormat(annotation.dateFormat()).parse(String.valueOf(value)))
                        : value);
        }
        //string reading from excel is blank
        if (!toClass.isAssignableFrom(String.class)
                && "".equals(value))
        {
            return null;
        }

        return (T) value;
    }

    /**
     * Set excel title line's number(start from 0).Default is 0.
     *
     * @param titleLine start calculate from not null row
     * @return return object self
     */
    public PbPOIExcelTemplate setTitleLine(int titleLine)
    {
        this.titleLine = titleLine;
        return this;
    }

    /**
     * Set content to cells, based on the type of content
     *
     * @param cell            the cell
     * @param content         the content object
     * @param pbPOIExcelTitle the annotation
     */
    private void set(Cell cell, Object content, PbPOIExcelTitle pbPOIExcelTitle)
    {
        String textValue = null;

        //If the content object is null
        if (content == null) textValue = "";

        //If the content object is boolean
        if (content instanceof Boolean)
        {
            cell.setCellValue((Boolean) content);
        }
        //If the content object is Date
        else if (content instanceof Date)
        {
            cell.setCellValue(((Date) content));

            Workbook workbook = cell.getSheet()
                                    .getWorkbook();
            CellStyle cellStyle = workbook.createCellStyle();
            cellStyle.setDataFormat(workbook.createDataFormat()
                                            .getFormat(pbPOIExcelTitle.dateFormat()));
            cell.setCellStyle(cellStyle);
        }
        //If the content is a picture
        else if (content instanceof byte[])
        {
        }
        else if (content instanceof File)
        {
        }
        //If the content is a picture file path
        else if ((content instanceof String)
                && pbPOIExcelTitle.isFilePath())
        {
        }
        //Other
        else
        {
            textValue = content != null ? content.toString() : null;
        }

        //If the content is a number or string
        if (textValue != null)
        {
            //If the content is a number
            Pattern p = Pattern.compile("^//d+(//.//d+)?$");
            Matcher matcher = p.matcher(textValue);
            if (matcher.matches())
            {
                cell.setCellValue(Double.parseDouble(textValue));
            }
            //If the content is a string
            else
            {
                cell.setCellValue(textValue);
            }
        }
    }
}