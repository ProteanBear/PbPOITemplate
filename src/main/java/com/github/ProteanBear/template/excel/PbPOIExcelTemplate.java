package com.github.ProteanBear.template.excel;

import com.github.ProteanBear.template.annotation.PbPOIExcel;
import com.github.ProteanBear.template.annotation.PbPOIExcelTitle;
import com.github.ProteanBear.template.utils.ClassUtils;
import com.github.ProteanBear.template.utils.FileSuffix;
import com.github.ProteanBear.template.utils.Hex26Utils;
import com.github.ProteanBear.template.utils.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * The tools is created for easy use of Apache POI.
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
    private int titleLine=0;

    /**
     * Read excel file and write to list of T class
     *
     * @param excelFile The file .xls or .xlsx
     * @return
     */
    public List<?> readFrom(File excelFile,Class<?> returnClass)
            throws IOException, InvalidFormatException, NoSuchMethodException, IllegalAccessException,
            InvocationTargetException, InstantiationException
    {
        //Load Excel File
        Workbook workbook=WorkbookFactory.create(excelFile);

        //Get returnClass title->setMethod map by annotation
        Map<String,Method> titleMethodMap=ClassUtils.titleMapSetMethodBy(PbPOIExcelTitle.class,returnClass);

        //Record index->title map
        Map<String,String> indexTitleMap=new HashMap<String,String>();

        //All sheets
        Sheet sheet=null;
        Row row=null;
        List<Object> result=new ArrayList<>();
        int pageNum=workbook.getNumberOfSheets();
        for(int page=0;page<pageNum;page++)
        {
            //Get current sheet
            sheet=workbook.getSheetAt(page);

            //All rows be not null
            Iterator<Row> rowItr=sheet.rowIterator();
            int rowNum=0;
            while(rowItr.hasNext())
            {
                //Get current row
                row=rowItr.next();
                int colNum=row.getLastCellNum();

                //Set index title for not title excel
                if(titleLine==-1 && rowNum==0)
                {
                    //Set title to "A……Z" mode
                    for(int i=0;i<colNum;i++)
                    {
                        indexTitleMap.put(i+"",Hex26Utils.from(i+1));
                    }
                }
                //If current row is title line
                //,record into index->title map
                else if(titleLine==rowNum)
                {
                    for(int index=0;index<colNum;index++)
                    {
                        indexTitleMap.put(index+"",valueOf(row.getCell(index))+"");
                    }
                    rowNum++;
                    continue;
                }

                //Create new tClass Object instance
                Object object=returnClass.getDeclaredConstructor().newInstance();
                //All cells include null cell
                for(int index=0;index<colNum;index++)
                {
                    //Set field content,index->title->setMethod
                    Method method=titleMethodMap.get(indexTitleMap.get(index+""));
                    if(method==null) continue;
                    method.invoke(object,valueOf(row.getCell(index)));
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
     */
    public void writeTo(File excelFile,List<?>... data) throws IOException
    {
        //File exists, delete the old file
        if(excelFile.exists()) excelFile.delete();
        excelFile.createNewFile();

        writeTo(FileSuffix.getBy(excelFile),
                new FileOutputStream(excelFile),
                data
        );
    }

    /**
     * Writes the specified data list to an Excel file.
     *
     * @param outputStream Output stream
     * @param data         the specified data list.Multiple sets of data to generate multiple sheets.
     */
    public void writeTo(FileSuffix fileSuffix,OutputStream outputStream,List<?>... data) throws IOException
    {
        switch(fileSuffix)
        {
            case EXCEL_XLSX:
                writeToXLSX(data).write(outputStream);
                break;
            case EXCEL_XLS:
                writeToXLS(data).write(outputStream);
                break;
            default:
        }
    }

    /**
     * return cell's value object
     *
     * @param cell Excel
     * @return Object,eg. String/Double/Date/Boolean/null
     */
    public Object valueOf(Cell cell)
    {
        Object result=null;

        //Get cell's type
        CellType type=cell.getCellTypeEnum();
        //handle different type
        switch(type)
        {
            //String
            case STRING:
            {
                result=cell.getRichStringCellValue().toString().trim();
                break;
            }
            //Number(Double)
            case NUMERIC:
                //Formula(Double)
            case FORMULA:
            {
                //Date check
                //Return java.util.Date object
                if(org.apache.poi.ss.usermodel.DateUtil.isCellDateFormatted(cell))
                {
                    result=cell.getDateCellValue();
                }
                //Normal Number
                //Return Double object
                else
                {
                    result=cell.getNumericCellValue();
                }
                break;
            }
            //Boolean
            case BOOLEAN:
            {
                result=cell.getBooleanCellValue();
                break;
            }
            //其他
            default:
        }

        return result;
    }

    /**
     * Set excel title line's number(start from 0).Default is 0.
     *
     * @param titleLine start calculate from not null row
     * @return return object self
     */
    public PbPOIExcelTemplate setTitleLine(int titleLine)
    {
        this.titleLine=titleLine;
        return this;
    }

    /**
     * Generate the excel content by the data.
     *
     * @param data the specified data list.Multiple sets of data to generate multiple sheets.
     */
    protected XSSFWorkbook writeToXLSX(List<?>... data) throws IOException
    {
        //Declare a workbook
        XSSFWorkbook workbook=new XSSFWorkbook();

        //Create the sheets
        XSSFSheet sheet;
        XSSFRow row;
        XSSFCell cell;
        Class curClass;
        PbPOIExcel pbPOIExcelAnnotation;
        Map<PbPOIExcelTitle,Method> getMethodMap;
        for(List<?> oneDataList : data)
        {
            if(oneDataList.isEmpty()) continue;

            //Get the class corresponding annotation
            curClass=oneDataList.get(0).getClass();
            pbPOIExcelAnnotation=oneDataList.get(0).getClass().getAnnotation(PbPOIExcel.class);
            if(pbPOIExcelAnnotation==null) continue;

            //Create a sheet
            String sheetTitle=pbPOIExcelAnnotation.sheetTitle();
            sheet=StringUtils.isBlank(sheetTitle)?workbook.createSheet():workbook.createSheet(sheetTitle);
            int curRow=0, index=0;

            //Generate the table title line
            try
            {
                getMethodMap=ClassUtils.titleMapGetMethodBy(PbPOIExcelTitle.class,curClass);

                row=sheet.createRow(curRow++);
                for(PbPOIExcelTitle pbPOIExcelTitle : getMethodMap.keySet())
                {
                    cell=row.createCell(index);
                    set(cell,pbPOIExcelTitle.value(),pbPOIExcelTitle);
                    index++;
                }
            }
            catch(Exception e)
            {
                e.printStackTrace();
                continue;
            }

            //Generate the table content line
            for(int i=0;i<oneDataList.size();i++)
            {
                row=sheet.createRow(curRow++);
                index=0;
                Object oneData=oneDataList.get(0);

                for(PbPOIExcelTitle pbPOIExcelTitle : getMethodMap.keySet())
                {
                    cell=row.createCell(index);
                    try
                    {
                        set(
                                cell,
                                getMethodMap.get(pbPOIExcelTitle).invoke(oneData),
                                pbPOIExcelTitle
                        );
                    }
                    catch(Exception e)
                    {
                        e.printStackTrace();
                        set(cell,"",pbPOIExcelTitle);
                    }
                    index++;
                }
            }
        }

        return workbook;
    }

    /**
     * Generate the excel content by the data.
     *
     * @param data the specified data list.Multiple sets of data to generate multiple sheets.
     */
    protected HSSFWorkbook writeToXLS(List<?>... data) throws IOException
    {
        //Declare a workbook
        HSSFWorkbook workbook=new HSSFWorkbook();

        //Create the sheets
        HSSFSheet sheet;
        HSSFRow row;
        HSSFCell cell;
        Class curClass;
        PbPOIExcel pbPOIExcelAnnotation;
        Map<PbPOIExcelTitle,Method> getMethodMap;
        for(List<?> oneDataList : data)
        {
            if(oneDataList.isEmpty()) continue;

            //Get the class corresponding annotation
            curClass=oneDataList.get(0).getClass();
            pbPOIExcelAnnotation=oneDataList.get(0).getClass().getAnnotation(PbPOIExcel.class);
            if(pbPOIExcelAnnotation==null) continue;

            //Create a sheet
            String sheetTitle=pbPOIExcelAnnotation.sheetTitle();
            sheet=StringUtils.isBlank(sheetTitle)?workbook.createSheet():workbook.createSheet(sheetTitle);
            int curRow=0, index=0;

            //Generate the table title line
            try
            {
                getMethodMap=ClassUtils.titleMapGetMethodBy(PbPOIExcelTitle.class,curClass);

                row=sheet.createRow(curRow++);
                for(PbPOIExcelTitle pbPOIExcelTitle : getMethodMap.keySet())
                {
                    cell=row.createCell(index);
                    set(cell,pbPOIExcelTitle.value(),pbPOIExcelTitle);
                    index++;
                }
            }
            catch(Exception e)
            {
                e.printStackTrace();
                continue;
            }

            //Generate the table content line
            for(int i=0;i<oneDataList.size();i++)
            {
                row=sheet.createRow(curRow++);
                index=0;
                Object oneData=oneDataList.get(0);

                for(PbPOIExcelTitle pbPOIExcelTitle : getMethodMap.keySet())
                {
                    cell=row.createCell(index);
                    try
                    {
                        set(
                                cell,
                                getMethodMap.get(pbPOIExcelTitle).invoke(oneData),
                                pbPOIExcelTitle
                        );
                    }
                    catch(Exception e)
                    {
                        e.printStackTrace();
                        set(cell,"",pbPOIExcelTitle);
                    }
                    index++;
                }
            }
        }

        return workbook;
    }

    /**
     * Set content to cells, based on the type of content
     *
     * @param cell            the cell
     * @param content         the content object
     * @param pbPOIExcelTitle the annotation
     */
    protected void set(Cell cell,Object content,PbPOIExcelTitle pbPOIExcelTitle)
    {
        String textValue=null;

        //If the content object is null
        if(content==null) textValue="";

        //If the content object is boolean
        if(content instanceof Boolean)
        {
            cell.setCellType(CellType.BOOLEAN);
            cell.setCellValue(((Boolean)content).booleanValue());
        }
        //If the content object is Date
        else if(content instanceof Date)
        {
            cell.setCellType(CellType.NUMERIC);
            cell.setCellValue(((Date)content));

            Workbook workbook=cell.getSheet().getWorkbook();
            CellStyle cellStyle=workbook.createCellStyle();
            cellStyle.setDataFormat(workbook.createDataFormat().getFormat(pbPOIExcelTitle.dateFormat()));
            cell.setCellStyle(cellStyle);
        }
        //If the content is a picture
        else if(content instanceof byte[])
        {
        }
        else if(content instanceof File)
        {
        }
        //If the content is a picture file path
        else if((content instanceof String)
                && pbPOIExcelTitle.isFilePath())
        {
        }
        //Other
        else
        {
            textValue=content.toString();
        }

        //If the content is a number or string
        if(textValue!=null)
        {
            //If the content is a number
            Pattern p=Pattern.compile("^//d+(//.//d+)?$");
            Matcher matcher=p.matcher(textValue);
            if(matcher.matches())
            {
                cell.setCellType(CellType.NUMERIC);
                cell.setCellValue(Double.parseDouble(textValue));
            }
            //If the content is a string
            else
            {
                cell.setCellType(CellType.STRING);
                XSSFRichTextString richString=new XSSFRichTextString(
                        textValue);
                cell.setCellValue(richString);
            }
        }
    }
}