package com.github.ProteanBear.template.excel;

import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.junit.Assert;
import org.junit.Test;

import java.io.File;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.util.List;

/**
 * Test PbPOIExcelTemplate class
 *
 * @author ProteanBear
 */
public class PbPOIExcelTemplateTest
{
    /**
     * Test for loading excel file
     */
    @Test
    public void readExcelFile()
    {
        try
        {
            PbPOIExcelTemplate excelTemplate=new PbPOIExcelTemplate();
            List<ExcelTestBean> list=(List<ExcelTestBean>)excelTemplate.readFrom(new File(getClass().getResource("/normalWithTitle.xlsx").getPath()),ExcelTestBean.class);

            Assert.assertTrue("读取Excel错误！",(list!=null&&!list.isEmpty()));
            ObjectMapper objectMapper=new ObjectMapper();
            System.out.println(objectMapper.writerWithDefaultPrettyPrinter().writeValueAsString(list));
        }
        catch(NoSuchMethodException e)
        {
            e.printStackTrace();
        }
        catch(IllegalAccessException e)
        {
            e.printStackTrace();
        }
        catch(InvocationTargetException e)
        {
            e.printStackTrace();
        }
        catch(InvalidFormatException e)
        {
            e.printStackTrace();
        }
        catch(IOException e)
        {
            e.printStackTrace();
        }
        catch(InstantiationException e)
        {
            e.printStackTrace();
        }
    }
}