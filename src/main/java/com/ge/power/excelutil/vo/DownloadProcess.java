package com.ge.power.excelutil.vo;

import java.util.List;

public class DownloadProcess
{
    private List<ParamRefDtls> paramRefDtls;

    private List<IOData> iOData;

    private List<ProcessDtls> processDtls;
    
    private String ssoID;

	public List<ParamRefDtls> getParamRefDtls() {
		return paramRefDtls;
	}

	public void setParamRefDtls(List<ParamRefDtls> paramRefDtls) {
		this.paramRefDtls = paramRefDtls;
	}

	public List<IOData> getiOData() {
		return iOData;
	}

	public void setiOData(List<IOData> iOData) {
		this.iOData = iOData;
	}

	public List<ProcessDtls> getProcessDtls() {
		return processDtls;
	}

	public void setProcessDtls(List<ProcessDtls> processDtls) {
		this.processDtls = processDtls;
	}

	/**
	 * @return the ssoID
	 */
	public String getSsoID() {
		return ssoID;
	}

	/**
	 * @param ssoID the ssoID to set
	 */
	public void setSsoID(String ssoID) {
		this.ssoID = ssoID;
	}

    

}
