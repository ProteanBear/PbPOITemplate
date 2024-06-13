package xyz.proteanbear.template.word;

import org.apache.poi.common.usermodel.PictureType;
import xyz.proteanbear.template.annotation.PbPOIWordVariable;

import java.util.List;

public class WordTestDataBean
{
    @PbPOIWordVariable("case-number")
    private String caseNumber;

    @PbPOIWordVariable("organization-name")
    private String name;

    @PbPOIWordVariable("organization-specialties")
    private String specialties;

    @PbPOIWordVariable(value = "organization-image", isImagePath = true, imageType = PictureType.PNG, imageDescription = "测试图片", imageWidth = 166, imageHeight = 373)
    private String image;

    @PbPOIWordVariable(value = "test-image", isImagePath = true, imageType = PictureType.PNG, imageDescription = "测试图片2", imageWidth = 166, imageHeight = 373)
    private String testImage;

    @PbPOIWordVariable("doctor")
    private List<WordTestDoctorBean> doctorBeanList;

    @PbPOIWordVariable("drug")
    private List<WordTestDrugBean> drugBeanList;

    public List<WordTestDrugBean> getDrugBeanList()
    {
        return drugBeanList;
    }

    public void setDrugBeanList(List<WordTestDrugBean> drugBeanList)
    {
        this.drugBeanList=drugBeanList;
    }

    public String getCaseNumber()
    {
        return caseNumber;
    }

    public void setCaseNumber(String caseNumber)
    {
        this.caseNumber=caseNumber;
    }

    public String getName()
    {
        return name;
    }

    public void setName(String name)
    {
        this.name=name;
    }

    public String getSpecialties()
    {
        return specialties;
    }

    public void setSpecialties(String specialties)
    {
        this.specialties=specialties;
    }

    public List<WordTestDoctorBean> getDoctorBeanList()
    {
        return doctorBeanList;
    }

    public void setDoctorBeanList(List<WordTestDoctorBean> doctorBeanList)
    {
        this.doctorBeanList=doctorBeanList;
    }

    public String getImage()
    {
        return image;
    }

    public void setImage(String image)
    {
        this.image = image;
    }

    public String getTestImage()
    {
        return testImage;
    }

    public void setTestImage(String testImage)
    {
        this.testImage = testImage;
    }
}
