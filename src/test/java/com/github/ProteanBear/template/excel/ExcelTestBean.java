package com.github.ProteanBear.template.excel;

import com.fasterxml.jackson.annotation.JsonFormat;
import com.github.ProteanBear.template.annotation.PbPOIExcelTitle;

import java.util.Date;

/**
 *
 */
public class ExcelTestBean
{
    @PbPOIExcelTitle("标题")
    private String title;

    @PbPOIExcelTitle("总数")
    private Double count;

    @PbPOIExcelTitle("平均数")
    private Double average;

    @PbPOIExcelTitle("时间")
    @JsonFormat(pattern = "yyyy-MM-dd HH:mm",timezone = "GMT+8")
    private Date dateTime;

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
}