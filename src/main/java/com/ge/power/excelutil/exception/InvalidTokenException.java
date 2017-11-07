package com.ge.power.excelutil.exception;

import java.util.ArrayList;
import java.util.List;

public class InvalidTokenException extends RuntimeException{
	
	private static final long serialVersionUID = 1L;
	
	private List<String> errorMessages = new ArrayList<>();
	
	public InvalidTokenException(){}
	
	public InvalidTokenException(String msg){
		super(msg);
	}
	
	public InvalidTokenException(Throwable t){
		super(t);
	}

	public List<String> getErrorMessages(){
		return errorMessages;
	}

	public void addErrorMessages(String msg){
		this.errorMessages.add(msg);
	}
}

