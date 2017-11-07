package com.ge.power.excelutil.exception;

public class WebsilonDBTrasactionException extends Exception {
	public WebsilonDBTrasactionException(String msg) {
		super(msg);
		}
	
	private static final long serialVersionUID = 0xbdeb49c6d2241b5aL;

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
	public WebsilonDBTrasactionException(String exceptionCode, String exceptionMessage, String exceptionDetails) {
		super();
		this.exceptionCode = exceptionCode;
		this.exceptionMessage = exceptionMessage;
		this.exceptionDetails = exceptionDetails;
	}
}
