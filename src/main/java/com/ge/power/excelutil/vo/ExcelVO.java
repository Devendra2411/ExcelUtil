package com.ge.power.excelutil.vo;

import java.util.List;


public class ExcelVO {
	private String caseID;
	private String modelID;
	private List<ParamterVO> iOData;
	private String ssoID;
	public String getCaseID() {
		return caseID;
	}
	public void setCaseID(String caseID) {
		this.caseID = caseID;
	}
	public String getModelID() {
		return modelID;
	}
	public void setModelID(String modelID) {
		this.modelID = modelID;
	}
	public List<ParamterVO> getiOData() {
		return iOData;
	}
	public void setiOData(List<ParamterVO> iOData) {
		this.iOData = iOData;
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
