package xyz.proteanbear.template.word;

import xyz.proteanbear.template.annotation.PbPOIWordVariable;

public class WordTestDoctorBean
{
    @PbPOIWordVariable("name")
    private String name;
    @PbPOIWordVariable("practiceClass")
    private String practiceClass;
    @PbPOIWordVariable("practiceScope")
    private String practiceScope;
    @PbPOIWordVariable("qualify")
    private String qualify;
    @PbPOIWordVariable("level")
    private String level;
    @PbPOIWordVariable("year")
    private String year;
    @PbPOIWordVariable("is")
    private String is;

    public WordTestDoctorBean(String name,String practiceClass,String practiceScope,String qualify,String level,
                              String year,String is)
    {
        this.name=name;
        this.practiceClass=practiceClass;
        this.practiceScope=practiceScope;
        this.qualify=qualify;
        this.level=level;
        this.year=year;
        this.is=is;
    }

    public String getName()
    {
        return name;
    }

    public void setName(String name)
    {
        this.name=name;
    }

    public String getPracticeClass()
    {
        return practiceClass;
    }

    public void setPracticeClass(String practiceClass)
    {
        this.practiceClass=practiceClass;
    }

    public String getPracticeScope()
    {
        return practiceScope;
    }

    public void setPracticeScope(String practiceScope)
    {
        this.practiceScope=practiceScope;
    }

    public String getQualify()
    {
        return qualify;
    }

    public void setQualify(String qualify)
    {
        this.qualify=qualify;
    }

    public String getLevel()
    {
        return level;
    }

    public void setLevel(String level)
    {
        this.level=level;
    }

    public String getYear()
    {
        return year;
    }

    public void setYear(String year)
    {
        this.year=year;
    }

    public String getIs()
    {
        return is;
    }

    public void setIs(String is)
    {
        this.is=is;
    }
}