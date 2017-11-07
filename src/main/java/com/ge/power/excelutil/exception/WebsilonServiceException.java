package com.ge.power.excelutil.exception;

import java.io.Serializable;

import com.ge.power.excelutil.util.WebsilonConstant;


public class WebsilonServiceException extends Exception implements Serializable {
	private static final long serialVersionUID = 1L;

	public WebsilonServiceException() {
		super();
		}
	
	private String exceptionCode;
	private String exceptionMessage;
	private String exceptionDetails;

	public String getExceptionCode() {
		return exceptionCode;
	}
	public void setExceptionCode(String exceptionCode) {
		this.exceptionCode = exceptionCode;
	}
	public String getExceptionMessage() {
		return exceptionMessage;
	}
	public void setExceptionMessage(String exceptionMessage) {
		this.exceptionMessage = exceptionMessage;
	}
	public String getExceptionDetails() {
		return exceptionDetails;
	}
	public void setExceptionDetails(String exceptionDetails) {
		this.exceptionDetails = exceptionDetails;
	}
	public WebsilonServiceException(String exceptionCode, String exceptionMessage, String exceptionDetails) {
		super();
		this.exceptionCode = exceptionCode;
		this.exceptionMessage = exceptionMessage;
		this.exceptionDetails = exceptionDetails;
	}
	@Override
	public String toString() {
		String userMgmtException = "{"+WebsilonConstant.LEFT_SLASH+"exceptionCode"+WebsilonConstant.LEFT_SLASH+":"+WebsilonConstant.LEFT_SLASH+ exceptionCode+WebsilonConstant.LEFT_SLASH+","+WebsilonConstant.LEFT_SLASH+"exceptionMessage"+WebsilonConstant.LEFT_SLASH+":"+WebsilonConstant.LEFT_SLASH+exceptionMessage+WebsilonConstant.LEFT_SLASH+","+WebsilonConstant.LEFT_SLASH+"exceptionDetails"+WebsilonConstant.LEFT_SLASH+":"+WebsilonConstant.LEFT_SLASH+exceptionDetails+WebsilonConstant.LEFT_SLASH+"}";
		return userMgmtException;
	}
}
