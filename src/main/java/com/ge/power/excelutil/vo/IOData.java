package com.ge.power.excelutil.vo;

public class IOData
{
    private String param_label;

    private String param_type;

    private String testCondition;

    private String refCondition;

    private String units;

    public String getParam_label ()
    {
        return param_label;
    }

    public void setParam_label (String param_label)
    {
        this.param_label = param_label;
    }

    public String getParam_type ()
    {
        return param_type;
    }

    public void setParam_type (String param_type)
    {
        this.param_type = param_type;
    }

    public String getTestCondition ()
    {
        return testCondition;
    }

    public void setTestCondition (String testCondition)
    {
        this.testCondition = testCondition;
    }

    public String getRefCondition ()
    {
        return refCondition;
    }

    public void setRefCondition (String refCondition)
    {
        this.refCondition = refCondition;
    }

    public String getUnits ()
    {
        return units;
    }

    public void setUnits (String units)
    {
        this.units = units;
    }
   
}