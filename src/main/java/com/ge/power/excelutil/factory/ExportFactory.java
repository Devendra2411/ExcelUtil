package com.ge.power.excelutil.factory;

import java.util.ArrayList;
import java.util.List;
import java.util.ResourceBundle;
import java.util.StringTokenizer;

import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Component;

import com.ge.power.excelutil.vo.ExcelVO;
import com.ge.power.excelutil.vo.ParamterVO;

@Component
public abstract class ExportFactory{
	ResourceBundle exportprops = ResourceBundle.getBundle("export");
	
	@Value("${export.file.fileName}")
	protected String fileName;
	
	@Value("${export.file.path}")	
	protected String filePath;
	
	
	
	public abstract String export(List<ParamterVO> ExportData) throws Exception;
	
	protected List<String> getHeaders(){
		List<String> headers = new ArrayList<String>();
		StringTokenizer strtoken = new StringTokenizer(exportprops.getString("export.file.headers"),",");
		while(strtoken.hasMoreTokens()){
			headers.add(strtoken.nextToken());
		}
		for(String s:headers){
			System.out.println(s);
		}
		return headers;
	}
	protected String getTabOneName(){
		return exportprops.getString("export.excel.file.sheets");
	}
	protected String getFileName(){
		return new String().concat(filePath).concat(fileName).concat(".xls");
		}

	public abstract String downloadExcel(ExcelVO excelVO) throws Exception;
}
