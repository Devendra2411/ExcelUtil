package com.ge.power.excelutil.exception;

import java.util.ArrayList;
import java.util.List;

public class UnauthorizedException extends RuntimeException{
	
	private static final long serialVersionUID = 1L;
	
	private List<String> errorMessages = new ArrayList<>();
	
	public UnauthorizedException(){}
	
	public UnauthorizedException(String msg){
		super(msg);
	}
	
	public UnauthorizedException(Throwable t){
		super(t);
	}

	public List<String> getErrorMessages(){
		return errorMessages;
	}

	public void addErrorMessages(String msg){
		this.errorMessages.add(msg);
	}
}

