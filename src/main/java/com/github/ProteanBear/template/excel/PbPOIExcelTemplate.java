package com.github.ProteanBear.template.excel;

import com.github.ProteanBear.template.annotation.PbPOIExcelTitle;
import com.github.ProteanBear.template.utils.ClassUtils;
import com.github.ProteanBear.template.utils.Hex26Utils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.*;

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
     * @param excelFile The file .xls or .xlsx
     * @return
     */
    public List<?> readFrom(File excelFile,Class<?> returnClass)
            throws IOException, InvalidFormatException, NoSuchMethodException, IllegalAccessException, InvocationTargetException, InstantiationException
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
                if(titleLine==-1&&rowNum==0)
                {
                    //Set title to "A……Z" mode
                    for(int i=0;i<colNum;i++)
                    {
                        indexTitleMap.put(i+"",Hex26Utils.from(i));
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
}