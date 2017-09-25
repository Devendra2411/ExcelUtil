package com.ge.power.excelutil.vo;

public class ParamRefDtls
{
    private String param_margin;

    private String param_factor;

    private String param_corrected;

    private String param_type_value;
    
    private String margin_sense;
    
    private String param_label;

    public String getParam_margin ()
    {
        return param_margin;
    }

    public void setParam_margin (String param_margin)
    {
        this.param_margin = param_margin;
    }

    public String getParam_factor ()
    {
        return param_factor;
    }

    public void setParam_factor (String param_factor)
    {
        this.param_factor = param_factor;
    }

    public String getParam_corrected ()
    {
        return param_corrected;
    }

    public void setParam_corrected (String param_corrected)
    {
        this.param_corrected = param_corrected;
    }

    public String getParam_type_value ()
    {
        return param_type_value;
    }

    public void setParam_type_value (String param_type_value)
    {
        this.param_type_value = param_type_value;
    }

	/**
	 * @return the margin_sense
	 */
	public String getMargin_sense() {
		return margin_sense;
	}

	/**
	 * @param margin_sense the margin_sense to set
	 */
	public void setMargin_sense(String margin_sense) {
		this.margin_sense = margin_sense;
	}

	/**
	 * @return the param_label
	 */
	public String getParam_label() {
		return param_label;
	}

	/**
	 * @param param_label the param_label to set
	 */
	public void setParam_label(String param_label) {
		this.param_label = param_label;
	}

}
