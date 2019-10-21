package xyz.proteanbear.template.word;

import org.junit.jupiter.api.Test;
import xyz.proteanbear.template.PbPOIWordTemplate;
import xyz.proteanbear.template.utils.FileSuffix;

import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

public class PbPOIWordTemplateTest
{
    @Test
    public void writeWordFile()
    {
        try
        {
            File template=new File(getClass().getResource("/").toURI().getPath(),"wordTemplate.docx");

            PbPOIWordTemplate wordTemplate=new PbPOIWordTemplate();
            WordTestDataBean testDataBean=new WordTestDataBean();
            testDataBean.setCaseNumber("2019-A2222");
            testDataBean.setName("高新医院");
            testDataBean.setSpecialties("内科，外科");
            List<WordTestDoctorBean> list=new ArrayList<>();
            list.add(new WordTestDoctorBean(
                    "测试1",
                    "医师",
                    "内科",
                    "哈哈",
                    "基本",
                    "2019",
                    "哈哈"
                 )
            );
            list.add(new WordTestDoctorBean(
                    "测试2",
                    "医师",
                    "内科",
                    "哈哈",
                    "基本",
                    "2019",
                    "哈哈"
                 )
            );
            list.add(new WordTestDoctorBean(
                    "测试3",
                    "医师",
                    "外科",
                    "哈哈",
                    "基本",
                    "2019",
                    "哈哈"
                 )
            );
            list.add(new WordTestDoctorBean(
                    "测试4",
                    "医师",
                    "反倒是咖啡店飞快的将就放得开",
                    "返回的司法鉴定放假后",
                    "对方的几号放假倒海翻江兑换积分",
                    "2019",
                    "地方地方恒大华府和大家好"
                 )
            );
            testDataBean.setDoctorBeanList(list);
            List<WordTestDrugBean> drugBeanList=new ArrayList<>();
            drugBeanList.add(new WordTestDrugBean("中联1","biaoz","hah"));
            drugBeanList.add(new WordTestDrugBean("中联2","biaoz","hah"));
            drugBeanList.add(new WordTestDrugBean("中联3","biaoz","hah"));
            testDataBean.setDrugBeanList(drugBeanList);

            wordTemplate.writeTo(
                    template,
                    FileSuffix.WORD_DOCX,
                    new FileOutputStream(new File(getClass().getResource("/").toURI().getPath(),"wordTest.docx")),
                    testDataBean);
        }
        catch(Exception e)
        {
            e.printStackTrace();
        }
    }
}