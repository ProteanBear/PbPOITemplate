package com.github.ProteanBear.template.excel;

import com.fasterxml.jackson.databind.ObjectMapper;
import org.junit.Assert;
import org.junit.Test;

import java.io.File;
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
    public void readExcelFileNormalWithTitle()
    {
        try
        {
            File excelFile=new File(getClass().getResource("/normalWithTitle.xlsx").getPath());
            PbPOIExcelTemplate excelTemplate=new PbPOIExcelTemplate();
            List<ExcelTestBean> list=(List<ExcelTestBean>)excelTemplate
                    .readFrom(excelFile,ExcelTestBean.class);

            Assert.assertTrue("读取Excel错误！",(list!=null&&!list.isEmpty()));
            ObjectMapper objectMapper=new ObjectMapper();
            System.out.println(objectMapper.writerWithDefaultPrettyPrinter().writeValueAsString(list));
        }
        catch(Exception e)
        {
            e.printStackTrace();
        }
    }

    /**
     * Test for loading excel file
     */
    @Test
    public void readExcelFileNormalNoTitle()
    {
        try
        {
            File excelFile=new File(getClass().getResource("/normalNoTitle.xlsx").getPath());
            PbPOIExcelTemplate excelTemplate=new PbPOIExcelTemplate();
            List<ExcelTestBean> list=(List<ExcelTestBean>)excelTemplate
                    .setTitleLine(-1)
                    .readFrom(excelFile,ExcelNoTitleTestBean.class);

            Assert.assertTrue("读取Excel错误！",(list!=null&&!list.isEmpty()));
            ObjectMapper objectMapper=new ObjectMapper();
            System.out.println(objectMapper.writerWithDefaultPrettyPrinter().writeValueAsString(list));
        }
        catch(Exception e)
        {
            e.printStackTrace();
        }
    }
}