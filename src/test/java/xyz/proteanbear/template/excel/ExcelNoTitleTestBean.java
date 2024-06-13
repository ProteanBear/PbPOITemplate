package xyz.proteanbear.template.excel;

import com.fasterxml.jackson.annotation.JsonFormat;
import xyz.proteanbear.template.annotation.PbPOIExcel;
import xyz.proteanbear.template.annotation.PbPOIExcelTitle;

import java.util.Date;

/**
 *
 */
@PbPOIExcel(titleLine = -1)
public class ExcelNoTitleTestBean
{
    @PbPOIExcelTitle("A")
    private String title;

    @PbPOIExcelTitle("B")
    private Double count;

    @PbPOIExcelTitle("C")
    private Double average;

    @PbPOIExcelTitle("D")
    @JsonFormat(pattern = "yyyy-MM-dd HH:mm",timezone = "GMT+8")
    private Date dateTime;

    @PbPOIExcelTitle("E")
    private String description;

    public String getTitle()
    {
        return title;
    }

    public void setTitle(String title)
    {
        this.title=title;
    }

    public Double getCount()
    {
        return count;
    }

    public void setCount(Double count)
    {
        this.count=count;
    }

    public Double getAverage()
    {
        return average;
    }

    public void setAverage(Double average)
    {
        this.average=average;
    }

    public Date getDateTime()
    {
        return dateTime;
    }

    public void setDateTime(Date dateTime)
    {
        this.dateTime=dateTime;
    }

    public String getDescription()
    {
        return description;
    }

    public void setDescription(String description)
    {
        this.description=description;
    }
}