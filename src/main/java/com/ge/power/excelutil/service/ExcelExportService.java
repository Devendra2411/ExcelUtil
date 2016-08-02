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
import com.ge.power.excelutil.vo.ExcelVO;
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

	public String downloadExcel(ExcelVO excelVO) throws Exception {

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
	    
	    
	  
		XSSFSheet inputsheet= excelWorkbook.createSheet("IODesignTable");
		Row headerRow = inputsheet.createRow(0);
		headerRow.createCell(0).setCellValue("EBS Case");
		headerRow.createCell(1).setCellValue("Parameter Name");
		headerRow.createCell(2).setCellValue("Parameter Type");
		headerRow.createCell(3).setCellValue("Common Name and I/O Label");
		headerRow.createCell(4).setCellValue("Unit Set");
		headerRow.createCell(5).setCellValue("Input_or_Output");
		headerRow.createCell(6).setCellValue("Display Order");
		headerRow.createCell(7).setCellValue("Ebsilon Model Parameter Type");
		headerRow.createCell(8).setCellValue("Ebsilon Model Parameter Name");
		headerRow.createCell(9).setCellValue("Max/Min Unit Set");
		headerRow.createCell(10).setCellValue("Min Value Allowed");
		headerRow.createCell(11).setCellValue("Max Value Allowed");
		headerRow.createCell(12).setCellValue("Active");
		headerRow.createCell(13).setCellValue("Mandatory Input");
		headerRow.createCell(14).setCellValue("Allowed Dropdown Values or Formula");
		headerRow.createCell(15).setCellValue("Display On");
		headerRow.createCell(16).setCellValue("I/O Data Page for Display");
		headerRow.createCell(17).setCellValue("MyAE Data Base Table Name");
		headerRow.createCell(18).setCellValue("Transfer Ebsilon name to Setup Sheet?");
		
		
	
		try {
			if(excelVO != null && excelVO.getiOData() != null && excelVO.getiOData().size()>0){
				inputsheet=fillDownloadSheetData(excelVO.getiOData(),inputsheet);
			}
			
			
			
			
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		//sheet.protectSheet("password");
		String fileName= "IOParameters.xls";
		FileOutputStream file = new FileOutputStream(new File(fileName));
		excelWorkbook.write(file);		
		file.close();
		return fileName;
	
	}
	private XSSFSheet fillDownloadSheetData(List<ParamterVO> getiOData, XSSFSheet sheet) {
		
		int headerrownum=0;
		Cell cell=null;
		for(ParamterVO row:getiOData){
			Row excelRow = sheet.createRow(++headerrownum);
			for(int i=0;i<19;i++){			
				if(i==0){
					cell=	excelRow.createCell(i);
					cell.setCellValue(row.getCaseName());	
				}
				if(i==1){
					cell=	excelRow.createCell(i);
					cell.setCellValue(row.getParaName());	
				}
				if(i==2){
					cell=	excelRow.createCell(i);
					cell.setCellValue(row.getParaType());	
				}
				if(i==3){
					cell=	excelRow.createCell(i);
					cell.setCellValue(row.getUnitSet());	
				}
				if(i==4){
					cell=	excelRow.createCell(i);
					cell.setCellValue(row.getDisOrder());	
				}
				if(i==5){
					cell=	excelRow.createCell(i);
					cell.setCellValue(row.getEbsparaType());	
				}
				if(i==6){
					cell=	excelRow.createCell(i);
					cell.setCellValue(row.getMinMaxUnitSet());	
				}
				if(i==7){
					cell=	excelRow.createCell(i);
					cell.setCellValue(row.getMinValue());	
				}
				if(i==8){
					cell=	excelRow.createCell(i);
					cell.setCellValue(row.getMaxValue());	
				}
				if(i==9){
					cell=	excelRow.createCell(i);
					cell.setCellValue(row.getActive());	
				}
				if(i==10){
					cell=	excelRow.createCell(i);
					cell.setCellValue(row.getMandatory());	
				}
				if(i==11){
					cell=	excelRow.createCell(i);
					cell.setCellValue(row.getAllowDropdown());	
				}
				if(i==12){
					cell=	excelRow.createCell(i);
					cell.setCellValue(row.getDisplayOn());	
				}
				if(i==13){
					cell=	excelRow.createCell(i);
					cell.setCellValue(row.getiOData());	
				}
				if(i==14){
					cell=	excelRow.createCell(i);
					cell.setCellValue(row.getTableName());	
				}
				if(i==15){
					cell=	excelRow.createCell(i);
					cell.setCellValue(row.getSheet());	
				}
				
				
				
			}
		}
		return sheet;
	}
	
	
	
}
