package xyz.proteanbear.template.word;

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
}
