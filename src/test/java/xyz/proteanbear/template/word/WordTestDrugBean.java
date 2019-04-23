package xyz.proteanbear.template.word;

import xyz.proteanbear.template.annotation.PbPOIWordVariable;

public class WordTestDrugBean
{
    @PbPOIWordVariable("kind")
    private String kind;
    @PbPOIWordVariable("standard")
    private String standard;
    @PbPOIWordVariable("level")
    private String level;

    public WordTestDrugBean(String kind,String standard,String level)
    {
        this.kind=kind;
        this.standard=standard;
        this.level=level;
    }

    public String getKind()
    {
        return kind;
    }

    public void setKind(String kind)
    {
        this.kind=kind;
    }

    public String getStandard()
    {
        return standard;
    }

    public void setStandard(String standard)
    {
        this.standard=standard;
    }

    public String getLevel()
    {
        return level;
    }

    public void setLevel(String level)
    {
        this.level=level;
    }
}