package com.ge.power.excelutil.service;

import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.usermodel.XSSFDataValidationHelper;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import com.ge.power.excelutil.factory.ExportFactory;
import com.ge.power.excelutil.vo.DataSpanVO;
import com.ge.power.excelutil.vo.DownloadProcess;
import com.ge.power.excelutil.vo.ExcelVO;
import com.ge.power.excelutil.vo.IOData;
import com.ge.power.excelutil.vo.IOValuesVO;
import com.ge.power.excelutil.vo.ParamDtlsVO;
import com.ge.power.excelutil.vo.ParamRefDtls;
import com.ge.power.excelutil.vo.ParamterVO;
import com.ge.power.excelutil.vo.ProcessDtls;

@Service
public class ExcelExportService extends ExportFactory {

	@Override
	public String export(List<ParamterVO> exportData) throws Exception {
		XSSFWorkbook excelWorkbook = new XSSFWorkbook();
		XSSFSheet inputsheet = excelWorkbook.createSheet("Input");
		XSSFSheet ouputsheet = excelWorkbook.createSheet("Output");
		
		CellStyle style = excelWorkbook.createCellStyle();
		Font font = excelWorkbook.createFont();
		font.setBoldweight(Font.BOLDWEIGHT_BOLD);
	    font.setColor(HSSFColor.LIGHT_ORANGE.index);
	    font.setFontHeightInPoints((short)13);
	    style.setFont(font);
	   	
	   	CellStyle hdrStyle = excelWorkbook.createCellStyle();
		hdrStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
		hdrStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
		Font hdrFont= excelWorkbook.createFont();
		hdrFont.setColor(HSSFColor.BLACK.index);
		hdrFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
	    hdrStyle.setFont(hdrFont);

		Row headerRow = inputsheet.createRow(2);
		headerRow.createCell(0).setCellValue("Common Name and I/O Label");
		inputsheet.addMergedRegion(new CellRangeAddress(2, 2, 0, 4));
		headerRow.createCell(5).setCellValue("Parameter Name");
		inputsheet.addMergedRegion(new CellRangeAddress(2, 2, 5, 8));
		headerRow.createCell(9).setCellValue("Ebsilon Model Parameter Name");
		inputsheet.addMergedRegion(new CellRangeAddress(2, 2, 9, 11));
		headerRow.createCell(12).setCellValue("Units");
		headerRow.createCell(13).setCellValue("Value");
		inputsheet.addMergedRegion(new CellRangeAddress(2, 2, 13, 14));
		//headerRow.createCell(5).setCellValue("Input/Output");
		//headerRow.createCell(9).setCellValue("Max Value Allowed");
		//headerRow.createCell(10).setCellValue("Min Value Allowed");
		
		for(int i=0;i<=14;i++){
			if(i==0 || i==5 || i==9 || i==12 || i==13 ){
			   CellUtil.setAlignment(headerRow.getCell(i), excelWorkbook, CellStyle.ALIGN_CENTER); 
			}else{
			   headerRow.createCell(i);
			}
		}
		
		for(int i=0;i<=14;i++){
			headerRow.getCell(i).setCellStyle(hdrStyle);
		}

		headerRow = ouputsheet.createRow(0);
		headerRow.createCell(0).setCellValue("These results are for reference only and are not guaranteed.");
		ouputsheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 6));
		for(int i=1;i<=6;i++){
			headerRow.createCell(i);
		}for(int i=0;i<=6;i++){
			headerRow.getCell(i).setCellStyle(style);
		}	
					
		headerRow = ouputsheet.createRow(2);
		headerRow.createCell(0).setCellValue("Common Name and I/O Label");
		ouputsheet.addMergedRegion(new CellRangeAddress(2, 2, 0, 4));
		headerRow.createCell(5).setCellValue("Parameter Name");
		ouputsheet.addMergedRegion(new CellRangeAddress(2, 2, 5, 8));
		headerRow.createCell(9).setCellValue("Ebsilon Model Parameter Name");
		ouputsheet.addMergedRegion(new CellRangeAddress(2, 2, 9, 11));
		headerRow.createCell(12).setCellValue("Units");
		headerRow.createCell(13).setCellValue("Value");
		ouputsheet.addMergedRegion(new CellRangeAddress(2, 2, 13, 14));
		// headerRow.createCell(5).setCellValue("Input/Output");
		// headerRow.createCell(5).setCellValue("Max Value Allowed");
		// headerRow.createCell(6).setCellValue("Min Value Allowed");
		
		for(int i=0;i<=14;i++){
			if(i==0 || i==5 || i==9 || i==12 || i==13 ){
			   CellUtil.setAlignment(headerRow.getCell(i), excelWorkbook, CellStyle.ALIGN_CENTER); 
			}else{
			   headerRow.createCell(i);
			}
		}
		
		for(int i=0;i<=14;i++){
			headerRow.getCell(i).setCellStyle(hdrStyle);
		}
		
		try {
			List<ParamterVO> inputData = new ArrayList<ParamterVO>();
			List<ParamterVO> ouputData = new ArrayList<ParamterVO>();

			for (ParamterVO paramterVO : exportData) {
				if (paramterVO.getParam_type().equalsIgnoreCase("I") && paramterVO.getDisplay().equalsIgnoreCase("Y")) {
					inputData.add(paramterVO);
				} else if (paramterVO.getParam_type().equalsIgnoreCase("O") && paramterVO.getDisplay().equalsIgnoreCase("Y")) {
					ouputData.add(paramterVO);
				}
			}
		    inputsheet = fillSheetData(inputData, inputsheet);
			ouputsheet = fillSheetData(ouputData, ouputsheet);
		} catch (Exception e) {
			e.printStackTrace();
		}

		String fileName = "Ebsilon.xlsx";
		FileOutputStream file = new FileOutputStream(new File(fileName));
		excelWorkbook.write(file);
		file.close();
		return fileName;
	}

	private XSSFSheet fillSheetData(List<ParamterVO> exportData, XSSFSheet sheet) {
		int headerrownum = 2;
		Cell cell = null;
		for (ParamterVO row : exportData) {
			Row excelRow = sheet.createRow(++headerrownum);
			for (int i = 0; i <=14; i++) {
				if (i == 0) {
					cell = excelRow.createCell(i);
					cell.setCellValue(row.getParam_label());
					sheet.addMergedRegion(new CellRangeAddress(headerrownum, headerrownum, 0, 4));
				}
				if (i == 5) {
					cell = excelRow.createCell(i);
					cell.setCellValue(row.getParam_value_type());
					sheet.addMergedRegion(new CellRangeAddress(headerrownum, headerrownum, 5, 8));
				}
				if (i == 9) {
					cell = excelRow.createCell(i);
					cell.setCellValue(row.getParam_name());
					sheet.addMergedRegion(new CellRangeAddress(headerrownum, headerrownum, 9, 11));
				}
				if (i == 12) {
					cell = excelRow.createCell(i);
					cell.setCellValue(row.getParam_unit());
				}
				if (i == 13) {
					cell = excelRow.createCell(i);
					cell.setCellValue(row.getParam_value());
					sheet.addMergedRegion(new CellRangeAddress(headerrownum, headerrownum, 13, 14));
				}
			}
		}
		return sheet;
	}

	public String downloadExcel(ExcelVO excelVO) throws Exception {
    	
		XSSFWorkbook excelWorkbook = new XSSFWorkbook();
		
		XSSFSheet inputsheet = excelWorkbook.createSheet("IODesignTable");
		
		CellStyle hdrStyle = excelWorkbook.createCellStyle();
		hdrStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
		hdrStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
		Font hdrFont= excelWorkbook.createFont();
		hdrFont.setColor(HSSFColor.BLACK.index);
		hdrFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
	    hdrStyle.setFont(hdrFont);
	    
		Row headerRow = inputsheet.createRow(0);
		headerRow = inputsheet.createRow(1);
		headerRow = inputsheet.createRow(2);
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
		
		for(int i=0;i<=18;i++){
			headerRow.getCell(i).setCellStyle(hdrStyle);
		}

		try {
			if (excelVO != null && excelVO.getiOData() != null && excelVO.getiOData().size() > 0) {
				inputsheet = fillDownloadSheetData(excelVO.getiOData(), inputsheet);
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	    String fileName = "IOParameters.xlsx";
		FileOutputStream file = new FileOutputStream(new File(fileName));
		excelWorkbook.write(file);
		file.close();
		return fileName;
	}

	private XSSFSheet fillDownloadSheetData(List<ParamterVO> getiOData, XSSFSheet sheet) {
		int headerrownum = 2;
		Cell cell = null;
		for (ParamterVO row : getiOData) {
			Row excelRow = sheet.createRow(++headerrownum);
			for (int i = 0; i < 19; i++) {
				if (i == 0) {
					cell = excelRow.createCell(i);
					cell.setCellValue(row.getCaseName());
				}
				if (i == 1) {
					cell = excelRow.createCell(i);
					cell.setCellValue(row.getParaName());
				}
				if (i == 2) {
					cell = excelRow.createCell(i);
					cell.setCellValue(row.getParaType());
				}
				if (i == 3) {
					cell = excelRow.createCell(i);
					cell.setCellValue(row.getParam_label());
				}
				if (i == 4) {
					cell = excelRow.createCell(i);
					cell.setCellValue(row.getUnitSet());
				}
				if (i == 5) {
					cell = excelRow.createCell(i);
					cell.setCellValue(row.getParam_type());
				}
				if (i == 6) {
					cell = excelRow.createCell(i);
					cell.setCellValue(row.getDisOrder());
				}
				if (i == 7) {
					cell = excelRow.createCell(i);
					cell.setCellValue(row.getEbsparaType());
				}
				if (i == 8) {
					cell = excelRow.createCell(i);
					cell.setCellValue(row.getParam_name());
				}
				if (i == 9) {
					cell = excelRow.createCell(i);
					cell.setCellValue(row.getMinMaxUnitSet());
				}
				if (i == 10) {
					cell = excelRow.createCell(i);
					cell.setCellValue(row.getMinValue());
				}
				if (i == 11) {
					cell = excelRow.createCell(i);
					cell.setCellValue(row.getMaxValue());
				}
				if (i == 12) {
					cell = excelRow.createCell(i);
					cell.setCellValue(row.getActive());
				}
				if (i == 13) {
					cell = excelRow.createCell(i);
					cell.setCellValue(row.getMandatory());
				}
				if (i == 14) {
					cell = excelRow.createCell(i);
					cell.setCellValue(row.getAllowDropdown());
				}
				if (i == 15) {
					cell = excelRow.createCell(i);
					cell.setCellValue(row.getDisplayOn());
				}
				if (i == 16) {
					cell = excelRow.createCell(i);
					cell.setCellValue(row.getiOData());
				}
				if (i == 17) {
					cell = excelRow.createCell(i);
					cell.setCellValue(row.getTableName());
				}
				if (i == 18) {
					cell = excelRow.createCell(i);
					cell.setCellValue(row.getSheet());
				}
			}
		}
		for (int i = 0; i < sheet.getRow(sheet.getLastRowNum()).getPhysicalNumberOfCells(); ++i) {
			sheet.setColumnWidth(i, (int) (sheet.getColumnWidth(i) * 3.2));
		}
		return sheet;
	}

	@Override
	public String downloadProcessPage(DownloadProcess downloadProcess) throws Exception {
		XSSFWorkbook excelWorkbook = new XSSFWorkbook();
		DataValidation dataValidation = null;
		DataValidationConstraint constraint = null;
		DataValidationHelper validationHelper = null;
		XSSFSheet inputsheet = excelWorkbook.createSheet("Process");
		try {
			int row = 0;
			if (downloadProcess != null) {
				
				CellStyle header = excelWorkbook.createCellStyle();
				Font headerF = excelWorkbook.createFont();
				headerF.setBoldweight(Font.BOLDWEIGHT_BOLD);
				headerF.setColor(HSSFColor.LIGHT_ORANGE.index);
				headerF.setFontHeightInPoints((short)13);
			    header.setFont(headerF);
			   				
				CellStyle style = excelWorkbook.createCellStyle();
				Font font = excelWorkbook.createFont();
				font.setBoldweight(Font.BOLDWEIGHT_BOLD);
			    font.setColor(HSSFColor.BLACK.index);
			    font.setFontHeightInPoints((short)13);
			    style.setFont(font);
			   			    			    
			    CellStyle hdrStyle = excelWorkbook.createCellStyle();
				Font hdrFont = excelWorkbook.createFont();
				hdrFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
				hdrFont.setColor(HSSFColor.BLACK.index);
				hdrFont.setFontHeightInPoints((short)11);
				hdrStyle.setFont(hdrFont);
								
				CellStyle unlockedCell = excelWorkbook.createCellStyle();
				unlockedCell.setLocked(false);
				
				DataFormat format = excelWorkbook.createDataFormat();
				CellStyle hideCell = excelWorkbook.createCellStyle();
				hideCell.setHidden(true);
				hideCell.setDataFormat(format.getFormat(";;;"));
								
				Row headerRow = inputsheet.createRow(row);
				
				headerRow.createCell(0).setCellValue("These results are for reference only and are not guaranteed.");
				inputsheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 6));
				for(int i=1;i<=6;i++){
					headerRow.createCell(i);
				}for(int i=0;i<=6;i++){
					headerRow.getCell(i).setCellStyle(header);
				}
			    
				row=row+2;
				headerRow = inputsheet.createRow(row);
				headerRow.createCell(0).setCellValue("TEST INFORMATION");
				headerRow.getCell(0).setCellStyle(style);
				
				row = row + 1;
				if (downloadProcess.getProcessDtls() != null) {
					List<ProcessDtls> processDtls = downloadProcess.getProcessDtls();
					System.out.println("processDtls length :" + processDtls.size());
					System.out.println(processDtls.get(0).getModelName());

					headerRow = inputsheet.createRow(row);
					headerRow.createCell(0).setCellValue("Model Name");
					headerRow.getCell(0).setCellStyle(hdrStyle);
					headerRow.createCell(1).setCellValue(processDtls.get(0).getModelName());
					headerRow.createCell(2).setCellValue("Case Name");
					headerRow.getCell(2).setCellStyle(hdrStyle);
					headerRow.createCell(3).setCellValue(processDtls.get(0).getCaseName());
					headerRow.createCell(4).setCellValue("Customers");
					headerRow.getCell(4).setCellStyle(hdrStyle);
					headerRow.createCell(5).setCellValue(processDtls.get(0).getCustomer());
					headerRow.createCell(6).setCellValue("Project");
					headerRow.getCell(6).setCellStyle(hdrStyle);
					headerRow.createCell(7).setCellValue(processDtls.get(0).getProject());

					row = row + 1;
					headerRow = inputsheet.createRow(row);
					headerRow.createCell(0).setCellValue("Test Unit");
					headerRow.getCell(0).setCellStyle(hdrStyle);
					headerRow.createCell(1).setCellValue(processDtls.get(0).getTest_unit());
					headerRow.createCell(2).setCellValue("Analysis Date");
					headerRow.getCell(2).setCellStyle(hdrStyle);
					headerRow.createCell(3).setCellValue(processDtls.get(0).getAnalysis_date());
					headerRow.createCell(4).setCellValue("Comments");
					headerRow.getCell(4).setCellStyle(hdrStyle);
					headerRow.createCell(5).setCellValue(processDtls.get(0).getComments());
					headerRow.createCell(6).setCellValue("Test Engineer");
					headerRow.getCell(6).setCellStyle(hdrStyle);
					headerRow.createCell(7).setCellValue(processDtls.get(0).getTest_engg());

					row = row + 1;
					headerRow = inputsheet.createRow(row);
					headerRow.createCell(0).setCellValue("Test Date\\Time");
					headerRow.getCell(0).setCellStyle(hdrStyle);
					String[] value=processDtls.get(0).getTest_time().split("\\s+");
					headerRow.createCell(1).setCellValue(value[0]);
					headerRow.createCell(2).setCellValue("Time Zone");
					headerRow.getCell(2).setCellStyle(hdrStyle);
					headerRow.createCell(3).setCellValue(value[1]);
					headerRow.createCell(4).setCellValue(processDtls.get(0).getUser_setup_id());
					
					headerRow.getCell(4).setCellStyle(hideCell);
					
					row = row + 2;
					headerRow = inputsheet.createRow(row);
					headerRow.createCell(0).setCellValue("UNIT SYSTEM:");
					headerRow.getCell(0).setCellStyle(hdrStyle);
					
					validationHelper=new XSSFDataValidationHelper(inputsheet);
				    CellRangeAddressList addressList = new  CellRangeAddressList(row,row,1,1);
				    constraint =validationHelper.createExplicitListConstraint(new String[]{"SI","ENG"});
				    dataValidation = validationHelper.createValidation(constraint, addressList);
				    dataValidation.setSuppressDropDownArrow(true);      
				    inputsheet.addValidationData(dataValidation);
				   // headerRow = inputsheet.createRow(8);
				    
				    inputsheet.getRow(7).createCell(1).setCellStyle(unlockedCell);
				  //  inputsheet.getRow(8).createCell(1).setCellStyle(unlockedCell);
				    if(processDtls.get(0).getUnitSystem().equalsIgnoreCase("SI"))
				    	inputsheet.getRow(row).getCell(1).setCellValue("SI");
				    else
				    	inputsheet.getRow(row).getCell(1).setCellValue("ENG");
				}

				if (downloadProcess.getiOData() != null) {
					List<IOData> inputData = new ArrayList<IOData>();
					List<IOData> outputData = new ArrayList<IOData>();

					for (IOData paramterVO : downloadProcess.getiOData()) {
						if (paramterVO.getParam_type().equalsIgnoreCase("I")) {
							inputData.add(paramterVO);
						} else if (paramterVO.getParam_type().equalsIgnoreCase("O")) {
							outputData.add(paramterVO);
						}
					}
					row = 9;
					headerRow = inputsheet.createRow(row);
					headerRow.createCell(0).setCellValue("OPERATING CONDITIONS");
					headerRow.getCell(0).setCellStyle(style);

					row = row + 1;
					headerRow = inputsheet.createRow(row);
					headerRow.createCell(0).setCellValue("INPUT PARAMETER");
					headerRow.createCell(1).setCellValue("UNITS");
					headerRow.createCell(2).setCellValue("TEST CONDITIONS");
					headerRow.createCell(3).setCellValue("GUARANTEE CONDITIONS");
					
					for(int i=0;i<4;i++){
						headerRow.getCell(i).setCellStyle(hdrStyle);
					}

					row = row + 1;
					for (IOData paramterVO : inputData) {
						headerRow = inputsheet.createRow(row);
						headerRow.createCell(0).setCellValue(paramterVO.getParam_label());
						headerRow.createCell(1).setCellValue(paramterVO.getUnits());
						headerRow.createCell(2).setCellValue(paramterVO.getTestCondition());
						headerRow.createCell(3).setCellValue(paramterVO.getRefCondition());
						headerRow.createCell(4).setCellValue("");
						headerRow.createCell(5).setCellValue(paramterVO.getParam_type());
						row++;
						
						for(int i=0;i<=5;i++){
							if(i == 2){
								headerRow.getCell(i).setCellStyle(unlockedCell);
							}else if(i==4 || i==5){
								headerRow.getCell(i).setCellStyle(hideCell);
							}
						}
					}
									

					row = row + 2;
					headerRow = inputsheet.createRow(row);
					headerRow.createCell(0).setCellValue("OPERATING RESULTS");
					headerRow.getCell(0).setCellStyle(style);
					
					row = row + 1;
					headerRow = inputsheet.createRow(row);
					headerRow.createCell(0).setCellValue("OUTPUT PARAMETER");
					headerRow.createCell(1).setCellValue("UNITS");
					headerRow.createCell(2).setCellValue("TEST RESULTS");
					headerRow.createCell(3).setCellValue("GUARANTEE");
					headerRow.createCell(4).setCellValue("MARGIN SENSE");
					headerRow.createCell(5);
					for(int i=0;i<=5;i++){
						headerRow.getCell(i).setCellStyle(hdrStyle);
					}	
					
					row = row + 1;
					for (IOData paramterVO : outputData) {
						
						headerRow = inputsheet.createRow(row);
						headerRow.createCell(0).setCellValue(paramterVO.getParam_label());
						headerRow.createCell(1).setCellValue(paramterVO.getUnits());
						headerRow.createCell(2).setCellValue(paramterVO.getTestCondition());
						headerRow.createCell(3).setCellValue(paramterVO.getRefCondition());
						headerRow.createCell(4);
						headerRow.createCell(5).setCellValue(paramterVO.getParam_type());
											
						validationHelper=new XSSFDataValidationHelper(inputsheet);
					    CellRangeAddressList addressList = new  CellRangeAddressList(row,row+1,4,4);
					    constraint =validationHelper.createExplicitListConstraint(new String[]{"Corrected to Reference","Reference to Corrected"});
					    dataValidation = validationHelper.createValidation(constraint, addressList);
					    dataValidation.setSuppressDropDownArrow(true);      
					    inputsheet.addValidationData(dataValidation);    
					    inputsheet.getRow(row).getCell(4).setCellStyle(unlockedCell);
					    if(paramterVO.getMargin_sense().equalsIgnoreCase("Corrected to Reference"))
					    	inputsheet.getRow(row).getCell(4).setCellValue("Corrected to Reference");	
					    else if(paramterVO.getMargin_sense().equalsIgnoreCase("Reference to Corrected"))
					    	inputsheet.getRow(row).getCell(4).setCellValue("Reference to Corrected");
					    inputsheet.createRow(row+1).createCell(4).setCellStyle(unlockedCell);
						
						for(int i=0;i<=5;i++){
								if(i==2){
									headerRow.getCell(i).setCellStyle(unlockedCell);
								}else if(i==5){
									headerRow.getCell(i).setCellStyle(hideCell);
								}
						}
						row++;
					}
				}

				row = row + 2;
				headerRow = inputsheet.createRow(row);
				headerRow.createCell(0).setCellValue("PERFORMANCE CORRECTIONS");
				headerRow.getCell(0).setCellStyle(style);
				
				row = row + 1;
				headerRow = inputsheet.createRow(row);
				headerRow.createCell(0).setCellValue("OUTPUT PARAMETERS");
				headerRow.createCell(1).setCellValue("FACTOR");
				headerRow.createCell(2).setCellValue("CORRECTED");
				headerRow.createCell(3).setCellValue("MARGIN(%)");
				headerRow.createCell(4).setCellValue("MARGIN SENSE");
				for(int i=0;i<=4;i++){
					headerRow.getCell(i).setCellStyle(hdrStyle);
				}
				
				row = row + 1;
				if (downloadProcess.getParamRefDtls() != null) {
					List<ParamRefDtls> paramRefDtls = downloadProcess.getParamRefDtls();
					System.out.println("paramRefDtls : " + paramRefDtls.size());
					for (ParamRefDtls paramterVO : downloadProcess.getParamRefDtls()) {
						headerRow = inputsheet.createRow(row);
						headerRow.createCell(0).setCellValue(paramterVO.getParam_label());
						headerRow.createCell(1).setCellValue(paramterVO.getParam_factor());
						headerRow.createCell(2).setCellValue(paramterVO.getParam_corrected());
						headerRow.createCell(3).setCellValue(paramterVO.getParam_margin());
						headerRow.createCell(4).setCellValue(paramterVO.getMargin_sense());
						headerRow.createCell(5).setCellValue("C_O");
						row++;
						
						for(int i=0;i<=5;i++){
							if(i==5){
								headerRow.getCell(i).setCellStyle(hideCell);
							}
						}
					}
				}
			}
			for (int i = 0; i < inputsheet.getRow(inputsheet.getFirstRowNum()+3).getPhysicalNumberOfCells(); ++i) {
				inputsheet.setColumnWidth(i, (int) (inputsheet.getColumnWidth(i) * 3));
			}
			
			inputsheet.enableLocking();
		} catch (Exception e) {
			e.printStackTrace();
		}
		String fileName = "Process.xlsx";
		//String fileName = "D://Gegdc//AS434512//Downloads//Process.xlsx";//for local
		FileOutputStream file = new FileOutputStream(new File(fileName));
		excelWorkbook.write(file);
		file.close();
		return fileName;
	}

	@Override
    public String downloadDataSpan(DataSpanVO data) throws Exception {
           XSSFWorkbook excelWorkbook = new XSSFWorkbook();
     
           CellStyle style = excelWorkbook.createCellStyle();
           Font font = excelWorkbook.createFont();
           font.setBoldweight(Font.BOLDWEIGHT_BOLD);
           font.setColor(HSSFColor.LIGHT_ORANGE.index);
           font.setFontHeightInPoints((short)13);
           style.setFont(font);
           
           CellStyle hdrStyle = excelWorkbook.createCellStyle();
           Font hdrFont = excelWorkbook.createFont();
           hdrFont.setColor(HSSFColor.BLACK.index);
           hdrFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
           hdrFont.setFontHeightInPoints((short) 13);
           hdrStyle.setFont(hdrFont);
           
           CellStyle subHdrStyle = excelWorkbook.createCellStyle();
           Font subHdrFont = excelWorkbook.createFont();
           subHdrFont.setColor(HSSFColor.BLACK.index);
           subHdrFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
           subHdrFont.setFontHeightInPoints((short) 11);
           subHdrStyle.setFont(subHdrFont);
           
           XSSFSheet inputsheet = excelWorkbook.createSheet("Data Span");
           try {
                  int row = 0;
                  if (data != null) {
                        Row headerRow = inputsheet.createRow(row);
                        headerRow.createCell(0).setCellValue("These results are for reference only and are not guaranteed.");
                        inputsheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 6));
                        for(int i=1;i<=6;i++){
                               headerRow.createCell(i);
                        }for(int i=0;i<=6;i++){
                               headerRow.getCell(i).setCellStyle(style);
                        }      
                        
                        String date = null;
                        String time = null;
                        if (data.getStartTime() != null && !data.getStartTime().equalsIgnoreCase("")) {
                               String[] parts = data.getStartTime().split("\\s", 2);
                               date = parts[0];
                               time = parts[1];
                        }

                        row = row + 3;
                        headerRow = inputsheet.createRow(row);
                        headerRow.createCell(0).setCellValue("Model Name");
                        headerRow.getCell(0).setCellStyle(subHdrStyle);
                        headerRow.createCell(1).setCellValue(data.getModelName());
                        headerRow.createCell(2).setCellValue("Case Name");
                        headerRow.getCell(2).setCellStyle(subHdrStyle);
                        headerRow.createCell(3).setCellValue(data.getCaseName());
                        headerRow.createCell(4).setCellValue("Date of Run");
                        headerRow.getCell(4).setCellStyle(subHdrStyle);
                        headerRow.createCell(5).setCellValue(date);
                        headerRow.createCell(6).setCellValue("Time of Run");
                        headerRow.getCell(6).setCellStyle(subHdrStyle);
                         headerRow.createCell(7).setCellValue(time);
                        headerRow.createCell(8).setCellValue("Run ID");
                        headerRow.getCell(8).setCellStyle(subHdrStyle);
                        headerRow.createCell(9).setCellValue(data.getRunID());
                        

                        if (data.getiOData() != null) {
                               List<ParamDtlsVO> inputData = new ArrayList<ParamDtlsVO>();
                               List<ParamDtlsVO> ouputData = new ArrayList<ParamDtlsVO>();

                               for (ParamDtlsVO paramterVO : data.getiOData()) {
                                      if ((paramterVO.getParam_type().equalsIgnoreCase("I1")
                                                    && paramterVO.getIteration().equalsIgnoreCase("1")) || (paramterVO.getParam_type().equalsIgnoreCase("I2")&& paramterVO.getIteration().equalsIgnoreCase("1"))) {
                                             inputData.add(paramterVO);
                                      } else if (paramterVO.getParam_type().equalsIgnoreCase("O_P")
                                                    && paramterVO.getIteration().equalsIgnoreCase("1")
                                                    && !paramterVO.getParam_label().equalsIgnoreCase("Program Return Status")) {
                                             ouputData.add(paramterVO);
                                      }
                               }
                               row = row + 2;
                               headerRow = inputsheet.createRow(row);
                               headerRow.createCell(0).setCellValue("Data Span Input");
                               headerRow.getCell(0).setCellStyle(hdrStyle);
                               row = row + 1;
                               headerRow = inputsheet.createRow(row);
                               headerRow.createCell(0).setCellValue("Input Parameter(s)");
                               headerRow.createCell(1).setCellValue("Initial Value");
                               headerRow.createCell(2).setCellValue("Final Value");
                               headerRow.createCell(3).setCellValue("Value Increment");
                               for (int i = 0; i <= 3; i++) {
                                      headerRow.getCell(i).setCellStyle(subHdrStyle);
                               }
                               row = row + 1;
                               for (ParamDtlsVO paramterVO : inputData) {
                                      headerRow = inputsheet.createRow(row);
                                      headerRow.createCell(0).setCellValue(paramterVO.getParam_label());
                                      headerRow.createCell(1).setCellValue(paramterVO.getInitialValue());
                                      headerRow.createCell(2).setCellValue(paramterVO.getFinalValue());
                                      headerRow.createCell(3).setCellValue(paramterVO.getIncreamentValue());
                                      row++;
                               }

                               /*headerRow = inputsheet.createRow(row);
                               headerRow.createCell(0).setCellValue("Output Parameter");
                               headerRow.getCell(0).setCellStyle(subHdrStyle);

                               row = row + 1;
                               for (ParamDtlsVO paramterVO : ouputData) {
                                      headerRow = inputsheet.createRow(row);
                                      headerRow.createCell(0).setCellValue(paramterVO.getParam_label());
                                      row++;
                               }*/
                        
                                                                  
                        /*     if (data.getTableViewVO() != null) {
                                      if (data.getTableViewVO().getxValue() != null) {
                                             for (int i = 1; i <= data.getTableViewVO().getxValue().size(); i++) {
                                                    String xValue = data.getTableViewVO().getxValue().get(i - 1);
                                                    headerRow.createCell(i).setCellValue(xValue);
                                                    headerRow.getCell(i).setCellStyle(subHdrStyle);
                                             }
                                      }
                                      row = row + 1;
                                      if (data.getTableViewVO().getyValue() != null) {
                                      for (java.util.Map.Entry<String, List<IOValuesVO>> entry : data.getTableViewVO().getyValue().entrySet()){
                                                    headerRow = inputsheet.createRow(row);
                                                    String key = entry.getKey();
                                                    if (key != null && !key.equalsIgnoreCase("")) {
                                                           headerRow.createCell(0).setCellValue(ouputData.get(0).getParam_label() + "("
                                                                        + inputData.get(1).getParam_label() + "=" + key + ")");
                                                           headerRow.getCell(0).setCellStyle(subHdrStyle);
                                                           List<IOValuesVO> list = entry.getValue();
                                                           for (int i = 1; i <= list.size(); i++) {
                                                                  IOValuesVO iOValuesVO = list.get(i - 1);
                                                                  headerRow.createCell(i);
                                                                  if (!iOValuesVO.getOutput().equalsIgnoreCase("empty")) {
                                                                         headerRow.getCell(i).setCellValue(iOValuesVO.getOutput());
                                                                  } else {
                                                                         headerRow.getCell(i).setCellValue("");
                                                                  }
                                                           }

                                                    } else {
                                                           headerRow.createCell(0).setCellValue(ouputData.get(0).getParam_label());
                                                           headerRow.getCell(0).setCellStyle(subHdrStyle);
                                                           List<IOValuesVO> list = entry.getValue();
                                                           for (int i = 1; i <= list.size(); i++) {
                                                                  IOValuesVO iOValuesVO = list.get(i - 1);
                                                                  headerRow.createCell(i);
                                                                  if (!iOValuesVO.getOutput().equalsIgnoreCase("empty")) {
                                                                         headerRow.getCell(i).setCellValue(iOValuesVO.getOutput());
                                                                  } else {
                                                                         headerRow.getCell(i).setCellValue("");
                                                                  }
                                                           }
                                                    }
                                                    row = row + 1;
                                             }
                                      }
                               }*/
                               if(data.getTableViewVO()!=null){

                                      if (data.getTableViewVO().size() > 0) {
                                             for (int i = 0; i < data.getTableViewVO().size(); i++) {
                                                    if (i == 0) {
                                                           row = row + 1;
                                                           headerRow = inputsheet.createRow(row);
                                                           System.out.println("rowNum"+row);
                                                           headerRow.createCell(0).setCellValue("Data Span Result");
                                                           headerRow.getCell(0).setCellStyle(hdrStyle);

                                                           row = row + 1;
                                                           System.out.println("rowNum"+row);
                                                           headerRow = inputsheet.createRow(row);
                                                           headerRow.createCell(0).setCellValue(inputData.get(0).getParam_label());
                                                           System.out.println("label::"+headerRow.getCell(0).getStringCellValue());
                                                           headerRow.getCell(0).setCellStyle(subHdrStyle);
                                                           headerRow.createCell(1).setCellValue(inputData.get(0).getParam_unit());
                                                           headerRow.getCell(1).setCellStyle(subHdrStyle);
                                                           if (data.getTableViewVO().get(i).getxValue() != null) {
                                                                  for (int j = 1; j <= data.getTableViewVO().get(i).getxValue().size(); j++) {
                                                                        String xValue = data.getTableViewVO().get(i).getxValue().get(j - 1);
                                                                         headerRow.createCell(j+1).setCellValue(xValue);
                                                                         System.out.println("xvalue"+xValue);
                                                                         headerRow.getCell(j+1).setCellStyle(subHdrStyle);
                                                                  }
                                                           }
                                                    }

                                                    
                                                    if (data.getTableViewVO().get(i).getyValue() != null) {
                                                           for (java.util.Map.Entry<String, List<IOValuesVO>> entry : data.getTableViewVO().get(i)
                                                                        .getyValue().entrySet()) {
                                                                  row = row + 1;
                                                                  headerRow = inputsheet.createRow(row);
                                                                  String key = entry.getKey();
                                                                  
                                                                  if (key != null && !key.equalsIgnoreCase("")) {
                                                                        
                                                                        headerRow = inputsheet.createRow(row);
                                                                         headerRow.createCell(0).setCellValue(entry.getValue().get(0).getParam_label()
                                                                                      + "(" + inputData.get(1).getParam_label() + "=" + key + ")");
                                                                         headerRow.getCell(0).setCellStyle(subHdrStyle);
                                                                         headerRow.createCell(1).setCellValue(entry.getValue().get(0).getParam_unit());
                                                                        List<IOValuesVO> list = entry.getValue();
                                                                         for (int j = 1; j <= list.size(); j++) {
                                                                               IOValuesVO iOValuesVO = list.get(j - 1);
                                                                               headerRow.createCell(j+1);
                                                                               if (!iOValuesVO.getOutput().equalsIgnoreCase("empty")) {
                                                                                      headerRow.getCell(j+1).setCellValue(iOValuesVO.getOutput());
                                                                               } else {
                                                                                      headerRow.getCell(j+1).setCellValue("");
                                                                               }
                                                                               
                                                                        }
                                                                        

                                                                  } else {
                                                                        System.out.println("rowNum--2"+row);
                                                                        List<IOValuesVO> list = entry.getValue();
                                                                         headerRow.createCell(0).setCellValue(list.get(0).getParam_label());
                                                                         headerRow.getCell(0).setCellStyle(subHdrStyle);
                                                                         headerRow.createCell(1).setCellValue(list.get(0).getParam_unit());
                                                                        for (int j = 1; j <= list.size(); j++) {
                                                                               IOValuesVO iOValuesVO = list.get(j - 1);
                                                                               
                                                                               headerRow.createCell(j+1);
                                                                               if (!iOValuesVO.getOutput().equalsIgnoreCase("empty")) {
                                                                                      headerRow.getCell(j+1).setCellValue(iOValuesVO.getOutput());
                                                                               } else {
                                                                                      headerRow.getCell(j+1).setCellValue("");
                                                                               }
                                                                        }
                                                                  }
                                                                  
                                                           }
                                                    }
                                             }

                                      }
                               
                               }
                               if (data.getCurveFit() != null) {
                                      if (data.getCurveFit().size() > 0) {
                                             for (int i = 0; i < data.getCurveFit().size(); i++) {
                                                    if (i == 0) {
                                                           row = row + 1;
                                                           headerRow = inputsheet.createRow(row);
                                                           headerRow.createCell(0).setCellValue("Data Span Result - Curve");
                                                           headerRow.getCell(0).setCellStyle(hdrStyle);

                                                           row = row + 1;
                                                           headerRow = inputsheet.createRow(row);
                                                           headerRow.createCell(0).setCellValue(inputData.get(0).getParam_label());
                                                           headerRow.getCell(0).setCellStyle(subHdrStyle);
                                                           headerRow.createCell(1).setCellValue(inputData.get(0).getParam_unit());
                                                           if (data.getCurveFit().get(i).getxValue() != null) {
                                                                  for (int j = 1; j <= data.getCurveFit().get(i).getxValue().size(); j++) {
                                                                        String xValue = data.getCurveFit().get(i).getxValue().get(j - 1);
                                                                         headerRow.createCell(j+1).setCellValue(xValue);
                                                                         headerRow.getCell(j+1).setCellStyle(subHdrStyle);
                                                                  }
                                                           }
                                                    }

                                                    row = row + 1;
                                                    if (data.getCurveFit().get(i).getyValue() != null) {
                                                           for (java.util.Map.Entry<String, List<IOValuesVO>> entry : data.getCurveFit().get(i)
                                                                        .getyValue().entrySet()) {
                                                                  headerRow = inputsheet.createRow(row);
                                                                  String key = entry.getKey();
                                                                  if (key != null && !key.equalsIgnoreCase("")) {
                                                                         headerRow.createCell(0).setCellValue(entry.getValue().get(0).getParam_label()
                                                                                      + "(" + inputData.get(1).getParam_label() + "=" + key + ")");
                                                                         headerRow.getCell(0).setCellStyle(subHdrStyle);
                                                                         headerRow.createCell(1).setCellValue(entry.getValue().get(0).getParam_unit());
                                                                        List<IOValuesVO> list = entry.getValue();
                                                                        for (int j = 1; j <= list.size(); j++) {
                                                                               IOValuesVO iOValuesVO = list.get(j - 1);
                                                                               headerRow.createCell(j+1);
                                                                               if (!iOValuesVO.getOutput().equalsIgnoreCase("empty")) {
                                                                                      headerRow.getCell(j+1).setCellValue(iOValuesVO.getOutput());
                                                                               } else {
                                                                                      headerRow.getCell(j+1).setCellValue("");
                                                                               }
                                                                        }

                                                                  } else {
                                                                         headerRow.createCell(0).setCellValue(entry.getValue().get(0).getParam_label());
                                                                         headerRow.getCell(0).setCellStyle(subHdrStyle);
                                                                         headerRow.createCell(1).setCellValue(entry.getValue().get(0).getParam_unit());
                                                                        List<IOValuesVO> list = entry.getValue();
                                                                        for (int j = 1; j <= list.size(); j++) {
                                                                               IOValuesVO iOValuesVO = list.get(j - 1);
                                                                               headerRow.createCell(j+1);
                                                                               if (!iOValuesVO.getOutput().equalsIgnoreCase("empty")) {
                                                                                      headerRow.getCell(j+1).setCellValue(iOValuesVO.getOutput());
                                                                               } else {
                                                                                      headerRow.getCell(j+1).setCellValue("");
                                                                               }
                                                                        }
                                                                 }
                                                                  
                                                           }
                                                    }
                                             }

                                      }
                                      row = row + 2;
                                      headerRow = inputsheet.createRow(row);
                                      headerRow.createCell(0).setCellValue("Curve Fit Equation(s):");
                                      headerRow.getCell(0).setCellStyle(hdrStyle);

                                      row = row + 1;
                                      for (int i = 0; i < data.getCurveFit().size(); i++) {
                                             headerRow = inputsheet.createRow(row);
                                             headerRow.createCell(0).setCellValue(data.getCurveFit().get(i).getEquation());
                                             inputsheet.addMergedRegion(
                                                           new CellRangeAddress(headerRow.getRowNum(), headerRow.getRowNum(), 0, 2));
                                             headerRow.createCell(1);
                                             headerRow.createCell(2);
                                             for (int j = 0; j <= 2; j++) {
                                                    headerRow.getCell(j).setCellStyle(subHdrStyle);
                                             }
                                             row = row + 1;
                                      }
                               }
                        }
                  }

                  for (int i = 0; i < inputsheet.getRow(inputsheet.getFirstRowNum()+3).getPhysicalNumberOfCells(); ++i) {
                                      inputsheet.setColumnWidth(i, (int) (inputsheet.getColumnWidth(i) * 5));
                  }
           } catch (Exception e) {
                  e.printStackTrace();
           }

           String fileName = "DataSpan.xlsx";
           FileOutputStream file = new FileOutputStream(new File(fileName));
           excelWorkbook.write(file);
           file.close();
           return fileName;

    }


}
