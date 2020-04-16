package xyz.proteanbear.template.excel;

import com.fasterxml.jackson.databind.ObjectMapper;
import org.junit.jupiter.api.Test;
import xyz.proteanbear.template.PbPOIExcelTemplate;
import xyz.proteanbear.template.exception.FileSuffixNotSupportException;

import java.io.File;
import java.io.IOException;
import java.net.URISyntaxException;
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
            File excelFile = new File(getClass().getResource("/normalWithTitle.xlsx")
                                                .toURI()
                                                .getPath());
            PbPOIExcelTemplate excelTemplate = new PbPOIExcelTemplate();
            List<ExcelTestBean> list = excelTemplate
                    .readFrom(excelFile, ExcelTestBean.class);

            ObjectMapper objectMapper = new ObjectMapper();
            System.out.println(objectMapper.writerWithDefaultPrettyPrinter()
                                           .writeValueAsString(list));
        }
        catch (Exception e)
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
            File excelFile = new File(getClass().getResource("/normalNoTitle.xlsx")
                                                .toURI()
                                                .getPath());
            PbPOIExcelTemplate excelTemplate = new PbPOIExcelTemplate();
            List<ExcelNoTitleTestBean> list = excelTemplate
                    .setTitleLine(-1)
                    .readFrom(excelFile, ExcelNoTitleTestBean.class);

            ObjectMapper objectMapper = new ObjectMapper();
            System.out.println(objectMapper.writerWithDefaultPrettyPrinter()
                                           .writeValueAsString(list));
        }
        catch (Exception e)
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
        try
        {
            File excelFile = new File(getClass().getResource("/")
                                                .toURI()
                                                .getPath(), "writeTest.xlsx");
            List<ExcelTestBean> writeData = new ArrayList<>(5);
            for (int i = 0; i < 5; i++)
            {
                ExcelTestBean excelTestBean = new ExcelTestBean();
                excelTestBean.setTitle("标题内容" + i + "（" + (i % 2 == 1 ? "我长点" : "短") + "）");
                excelTestBean.setCount((double) i);
                excelTestBean.setAverage((double) i / 10.0);
                excelTestBean.setDateTime(new Date());
                writeData.add(excelTestBean);
            }

            PbPOIExcelTemplate excelTemplate = new PbPOIExcelTemplate();
            excelTemplate.writeTo(excelFile, writeData);
        }
        catch (IOException | URISyntaxException | FileSuffixNotSupportException e)
        {
            e.printStackTrace();
        }
    }
}