package xyz.proteanbear.template.excel;

import com.fasterxml.jackson.databind.ObjectMapper;
import org.junit.Assert;
import org.junit.Test;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
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
            File excelFile=new File(getClass().getResource("/1.xlsx").getPath());
            PbPOIExcelTemplate excelTemplate=new PbPOIExcelTemplate();
            List<ReportInfo> list=(List<ReportInfo>)excelTemplate
                    .readFrom(excelFile,ReportInfo.class);

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

    /**
     * Write a excel file.
     */
    @Test
    public void writeExcelFile()
    {
        File excelFile=new File(getClass().getResource("/").getPath(),"writeTest.xlsx");
        List<ExcelTestBean> writeData=new ArrayList<>(5);
        for(int i=0;i<5;i++)
        {
            ExcelTestBean excelTestBean=new ExcelTestBean();
            excelTestBean.setTitle("标题"+i);
            excelTestBean.setCount(new Double(i));
            excelTestBean.setAverage((new Double(i)).doubleValue()/10.0);
            excelTestBean.setDateTime(new Date());
            writeData.add(excelTestBean);
        }

        PbPOIExcelTemplate excelTemplate=new PbPOIExcelTemplate();
        try
        {
            excelTemplate.writeTo(excelFile,writeData);
        }
        catch(IOException e)
        {
            e.printStackTrace();
        }
    }
}