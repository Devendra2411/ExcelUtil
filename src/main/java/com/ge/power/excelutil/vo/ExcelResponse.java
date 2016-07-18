package com.ge.power.excelutil.vo;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelResponse {
	private String ouput;
	private XSSFWorkbook excelWorkbook;
	public String getOuput() {
		return ouput;
	}

	public void setOuput(String ouput) {
		this.ouput = ouput;
	}

	public XSSFWorkbook getExcelWorkbook() {
		return excelWorkbook;
	}

	public void setExcelWorkbook(XSSFWorkbook excelWorkbook) {
		this.excelWorkbook = excelWorkbook;
	}

}
