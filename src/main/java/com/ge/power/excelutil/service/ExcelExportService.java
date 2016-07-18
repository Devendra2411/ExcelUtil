package com.ge.power.excelutil.service;

import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import com.ge.power.excelutil.factory.ExportFactory;
import com.ge.power.excelutil.vo.ParamterVO;

@Service
public class ExcelExportService extends ExportFactory{

	@Override
	public String export(List<ParamterVO> exportData) throws Exception {
		List<String> headers = getHeaders();
		XSSFWorkbook excelWorkbook = new XSSFWorkbook();
		int colNum=0;
	    // Aqua background
		CellStyle editableCol = excelWorkbook.createCellStyle();
		editableCol.setFillBackgroundColor(IndexedColors.RED.getIndex());
		editableCol.setFillPattern(CellStyle.BIG_SPOTS);
		
        // Orange "foreground", foreground being the fill foreground not the font color.
		editableCol.setFillForegroundColor(IndexedColors.RED.getIndex());
		editableCol.setFillPattern(CellStyle.SOLID_FOREGROUND);
		Font font = excelWorkbook.createFont();
	    font.setFontHeightInPoints((short)10);
	    font.setFontName("GE Inspira");
	    font.setBoldweight(Font.BOLDWEIGHT_BOLD);
	    font.setColor(IndexedColors.WHITE.getIndex());
	    editableCol.setFont(font);
	    
	    
	    CellStyle noneditableCol = excelWorkbook.createCellStyle();
	    noneditableCol.setFillBackgroundColor(IndexedColors.GREEN.getIndex());
	    noneditableCol.setFillPattern(CellStyle.BIG_SPOTS);
		
        // Orange "foreground", foreground being the fill foreground not the font color.
		noneditableCol.setFillForegroundColor(IndexedColors.GREEN.getIndex());
		noneditableCol.setFillPattern(CellStyle.SOLID_FOREGROUND);
		font = excelWorkbook.createFont();
	    font.setFontHeightInPoints((short)10);
	    font.setFontName("GE Inspira");
	    font.setBoldweight(Font.BOLDWEIGHT_BOLD);
	    font.setColor(IndexedColors.WHITE.getIndex());
	    noneditableCol.setFont(font);
	    
	    
        //style.setWrapText(true);
        
        CellStyle style_data = excelWorkbook.createCellStyle();
	    Font font_data = excelWorkbook.createFont();
	    font_data.setFontHeightInPoints((short)10);
	    font_data.setFontName("GE Inspira");
	    
	    //For editable fields
	    
	    CellStyle unlockedCells = excelWorkbook.createCellStyle();
		font_data = excelWorkbook.createFont();
		font_data.setFontHeightInPoints((short)10);
		font_data.setFontName("GE Inspira");
		unlockedCells.setLocked(false);
	    
	    
	  
		XSSFSheet inputsheet= excelWorkbook.createSheet("input");
		XSSFSheet ouputsheet= excelWorkbook.createSheet("output");
		Row headerRow = inputsheet.createRow(0);
		headerRow.createCell(0).setCellValue("Common Name and I/O Label");
		headerRow.createCell(1).setCellValue("Parameter Name");
		headerRow.createCell(2).setCellValue("Ebsilon Name");
		headerRow.createCell(3).setCellValue("Units");
		headerRow.createCell(4).setCellValue("Value");
		//headerRow.createCell(5).setCellValue("Input/Output");
		headerRow.createCell(5).setCellValue("Max Value Allowed");
		headerRow.createCell(6).setCellValue("Min Value Allowed");
		
		headerRow = ouputsheet.createRow(0);
		headerRow.createCell(0).setCellValue("Common Name and I/O Label");
		headerRow.createCell(1).setCellValue("Parameter Name");
		headerRow.createCell(2).setCellValue("Ebsilon Name");
		headerRow.createCell(3).setCellValue("Units");
		headerRow.createCell(4).setCellValue("Value");
		//headerRow.createCell(5).setCellValue("Input/Output");
		headerRow.createCell(5).setCellValue("Max Value Allowed");
		headerRow.createCell(6).setCellValue("Min Value Allowed");
		/*for(String cellHeading : headers ){
			Cell cell=	headerRow.createCell(colNum++);
			if(colNum>15){
				cell.setCellStyle(noneditableCol);
			}else{cell.setCellStyle(editableCol);}
			cell.setCellValue(cellHeading);
		}*/
		try {
			List<ParamterVO> inputData = new ArrayList<ParamterVO>();
			List<ParamterVO> ouputData = new ArrayList<ParamterVO>();
			
			for (ParamterVO paramterVO : exportData) {
				if(paramterVO.getParam_type().equalsIgnoreCase("I")){
					inputData.add(paramterVO);
				}else if(paramterVO.getParam_type().equalsIgnoreCase("O")){
					ouputData.add(paramterVO);
				}
			}
			
			inputsheet=fillSheetData(inputData,inputsheet);
			ouputsheet=fillSheetData(ouputData,ouputsheet);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		//sheet.protectSheet("password");
		String fileName= "Ebsilon.xls";
		FileOutputStream file = new FileOutputStream(new File(fileName));
		excelWorkbook.write(file);		
		file.close();
		return fileName;
	}
	private XSSFSheet fillSheetData(List<ParamterVO> exportData,XSSFSheet sheet) {
		
		int headerrownum=0;
		Cell cell=null;
		for(ParamterVO row:exportData){
			Row excelRow = sheet.createRow(++headerrownum);
			for(int i=0;i<8;i++){			
				if(i==0){
					cell=	excelRow.createCell(i);
					cell.setCellValue(row.getParam_label());	
				}
				if(i==1){
					cell=	excelRow.createCell(i);
					cell.setCellValue("");	
				}
				if(i==2){
					cell=	excelRow.createCell(i);
					cell.setCellValue(row.getParam_name());	
				}
				if(i==3){
					cell=	excelRow.createCell(i);
					cell.setCellValue(row.getParam_unit());	
				}
				if(i==4){
					cell=	excelRow.createCell(i);
					cell.setCellValue(row.getParam_value());	
				}
				
			}
		}
		return sheet;
	}
	
	
	
}
