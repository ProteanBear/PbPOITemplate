package xyz.proteanbear.template.word;

import org.junit.jupiter.api.Test;
import xyz.proteanbear.template.PbPOIWordTemplate;
import xyz.proteanbear.template.utils.FileSuffix;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.net.URISyntaxException;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

public class PbPOIWordTemplateTest
{
    @Test
    public void writeWordFile()
            throws URISyntaxException, IOException, InvocationTargetException, NoSuchMethodException,
            IllegalAccessException
    {
        File template=new File(Objects.requireNonNull(getClass().getResource("/"))
                                      .toURI().getPath(), "wordTemplate.docx");

        PbPOIWordTemplate wordTemplate=new PbPOIWordTemplate();
        WordTestDataBean testDataBean=new WordTestDataBean();
        testDataBean.setCaseNumber("2019-A2222");
        testDataBean.setName("高新医院");
        testDataBean.setSpecialties("内科，外科");
        testDataBean.setImage("https://wsxk.care4u.cn/upload/recorder/record_flow.png");
        testDataBean.setTestImage("https://wsxk.care4u.cn/upload/recorder/license_flow.png");
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
                Files.newOutputStream(new File(Objects.requireNonNull(getClass().getResource("/"))
                                                      .toURI()
                                                      .getPath(), "wordTest.docx").toPath()),
                testDataBean);
    }
}