package com.ge.power.excelutil.vo;

import java.util.List;

public class DataSpanVO {
   
	private String modelName;
	private String modelID;
	private String actualModelName;
	private String caseName;
	private String caseID;
	private String ssoID;
	private String versionNo;
	
	private String userPrefID;
	private String userPrefName;
	private String universeFlag;
	private String active;
	
	private String createdBy;
	private String createdDate;
	private String updatedBy;
	private String updatedDate;
	
	private String status;
	private String statusMsg;
	
	private List<ParamDtlsVO> iOData;
	private List<ErrorVO> errorData;
	
	private TableViewVO tableViewVO;
	private List<TableViewVO> curveFit;
	
	private String runDate;
	private String runTime;
	private String runID;
	private String startTime;

	public String getModelName() {
		return modelName;
	}

	public void setModelName(String modelName) {
		this.modelName = modelName;
	}

	public String getModelID() {
		return modelID;
	}

	public void setModelID(String modelID) {
		this.modelID = modelID;
	}

	public String getActualModelName() {
		return actualModelName;
	}

	public void setActualModelName(String actualModelName) {
		this.actualModelName = actualModelName;
	}

	public String getCaseName() {
		return caseName;
	}

	public void setCaseName(String caseName) {
		this.caseName = caseName;
	}

	public String getCaseID() {
		return caseID;
	}

	public void setCaseID(String caseID) {
		this.caseID = caseID;
	}

	public String getSsoID() {
		return ssoID;
	}

	public void setSsoID(String ssoID) {
		this.ssoID = ssoID;
	}

	public String getVersionNo() {
		return versionNo;
	}

	public void setVersionNo(String versionNo) {
		this.versionNo = versionNo;
	}

	public String getUserPrefID() {
		return userPrefID;
	}

	public void setUserPrefID(String userPrefID) {
		this.userPrefID = userPrefID;
	}

	public String getUserPrefName() {
		return userPrefName;
	}

	public void setUserPrefName(String userPrefName) {
		this.userPrefName = userPrefName;
	}

	public String getUniverseFlag() {
		return universeFlag;
	}

	public void setUniverseFlag(String universeFlag) {
		this.universeFlag = universeFlag;
	}

	public String getActive() {
		return active;
	}

	public void setActive(String active) {
		this.active = active;
	}

	public String getCreatedBy() {
		return createdBy;
	}

	public void setCreatedBy(String createdBy) {
		this.createdBy = createdBy;
	}

	public String getCreatedDate() {
		return createdDate;
	}

	public void setCreatedDate(String createdDate) {
		this.createdDate = createdDate;
	}

	public String getUpdatedBy() {
		return updatedBy;
	}

	public void setUpdatedBy(String updatedBy) {
		this.updatedBy = updatedBy;
	}

	public String getUpdatedDate() {
		return updatedDate;
	}

	public void setUpdatedDate(String updatedDate) {
		this.updatedDate = updatedDate;
	}

	public String getStatus() {
		return status;
	}

	public void setStatus(String status) {
		this.status = status;
	}

	public String getStatusMsg() {
		return statusMsg;
	}

	public void setStatusMsg(String statusMsg) {
		this.statusMsg = statusMsg;
	}

	public List<ParamDtlsVO> getiOData() {
		return iOData;
	}

	public void setiOData(List<ParamDtlsVO> iOData) {
		this.iOData = iOData;
	}

	public List<ErrorVO> getErrorData() {
		return errorData;
	}

	public void setErrorData(List<ErrorVO> errorData) {
		this.errorData = errorData;
	}

	public TableViewVO getTableViewVO() {
		return tableViewVO;
	}

	public void setTableViewVO(TableViewVO tableViewVO) {
		this.tableViewVO = tableViewVO;
	}

	public String getRunDate() {
		return runDate;
	}

	public void setRunDate(String runDate) {
		this.runDate = runDate;
	}

	public String getRunTime() {
		return runTime;
	}

	public void setRunTime(String runTime) {
		this.runTime = runTime;
	}
	public String getRunID() {
		return runID;
	}

	public void setRunID(String runID) {
		this.runID = runID;
	}

	public List<TableViewVO> getCurveFit() {
		return curveFit;
	}

	public void setCurveFit(List<TableViewVO> curveFit) {
		this.curveFit = curveFit;
	}

	public String getStartTime() {
		return startTime;
	}

	public void setStartTime(String startTime) {
		this.startTime = startTime;
	}

}
