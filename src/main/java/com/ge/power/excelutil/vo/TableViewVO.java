package com.ge.power.excelutil.vo;

import java.util.List;
import java.util.Map;

public class TableViewVO {
	private List<String> xValue;
	private Map<String,List<IOValuesVO>> yValue;
	
	private String equation;
	
	public List<String> getxValue() {
		return xValue;
	}
	public void setxValue(List<String> xValue) {
		this.xValue = xValue;
	}
	public Map<String, List<IOValuesVO>> getyValue() {
		return yValue;
	}
	public void setyValue(Map<String, List<IOValuesVO>> yValue) {
		this.yValue = yValue;
	}
	public String getEquation() {
		return equation;
	}
	public void setEquation(String equation) {
		this.equation = equation;
	}

}
