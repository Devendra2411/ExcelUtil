package com.ge.power.excelutil.service;

import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
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
                    
                    
                  
                                XSSFSheet inputsheet= excelWorkbook.createSheet("Input");
                                XSSFSheet ouputsheet= excelWorkbook.createSheet("Output");
                                Row headerRow = inputsheet.createRow(0);
                                headerRow.createCell(0).setCellValue("Common Name and I/O Label");
                                inputsheet.addMergedRegion(new CellRangeAddress(0,0,0,4));
                                headerRow.createCell(5).setCellValue("Parameter Name");
                                inputsheet.addMergedRegion(new CellRangeAddress(0,0,5,6));
                                headerRow.createCell(7).setCellValue("Units");
                                headerRow.createCell(8).setCellValue("Value");
                                //headerRow.createCell(5).setCellValue("Input/Output");
                                headerRow.createCell(9).setCellValue("Max Value Allowed");
                                headerRow.createCell(10).setCellValue("Min Value Allowed");
                                
                                headerRow = ouputsheet.createRow(0);
                                headerRow.createCell(0).setCellValue("Common Name and I/O Label");
                                ouputsheet.addMergedRegion(new CellRangeAddress(0,0,0,4));
                                headerRow.createCell(5).setCellValue("Parameter Name");
                                ouputsheet.addMergedRegion(new CellRangeAddress(0,0,5,6));
                                headerRow.createCell(7).setCellValue("Units");
                                headerRow.createCell(8).setCellValue("Value");
                                //headerRow.createCell(5).setCellValue("Input/Output");
                                //headerRow.createCell(5).setCellValue("Max Value Allowed");
                                //headerRow.createCell(6).setCellValue("Min Value Allowed");
                                /*for(String cellHeading : headers ){
                                                Cell cell=              headerRow.createCell(colNum++);
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
                                String fileName= "D:\\gegdc\\DT476049\\Downloads\\Ebsilon.xlsx";
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
                                                for(int i=0;i<10;i++){                                       
                                                                if(i==0){
                                                                                cell=       excelRow.createCell(i);
                                                                                cell.setCellValue(row.getParam_label());
                                                                                sheet.addMergedRegion(new CellRangeAddress(headerrownum,headerrownum,0,4));
                                                                }
                                                                if(i==5){
                                                                                cell=       excelRow.createCell(i);
                                                                                cell.setCellValue(row.getParam_name());           
                                                                                sheet.addMergedRegion(new CellRangeAddress(headerrownum,headerrownum,5,6));
                                                                }
                                                                if(i==7){
                                                                                cell=       excelRow.createCell(i);
                                                                                cell.setCellValue(row.getParam_unit());               
                                                                }
                                                                if(i==8){
                                                                                cell=       excelRow.createCell(i);
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
                                
                                int headerrownum=2;
                                Cell cell=null;
                                for(ParamterVO row:getiOData){
                                                Row excelRow = sheet.createRow(++headerrownum);
                                                for(int i=0;i<19;i++){                                       
                                                                if(i==0){
                                                                                cell=       excelRow.createCell(i);
                                                                                cell.setCellValue(row.getCaseName()); 
                                                                }
                                                                if(i==1){
                                                                                cell=       excelRow.createCell(i);
                                                                                cell.setCellValue(row.getParaName()); 
                                                                }
                                                                if(i==2){
                                                                                cell=       excelRow.createCell(i);
                                                                                cell.setCellValue(row.getParaType());    
                                                                }
                                                                if(i==3){
                                                                                cell=       excelRow.createCell(i);
                                                                                cell.setCellValue(row.getParam_label());             
                                                                }
                                                                if(i==4){
                                                                                cell=       excelRow.createCell(i);
                                                                                cell.setCellValue(row.getUnitSet());       
                                                                }
                                                                if(i==5){
                                                                                cell=       excelRow.createCell(i);
                                                                                cell.setCellValue(row.getParam_type());              
                                                                }
                                                                if(i==6){
                                                                                cell=       excelRow.createCell(i);
                                                                                cell.setCellValue(row.getDisOrder());     
                                                                }
                                                                if(i==7){
                                                                                cell=       excelRow.createCell(i);
                                                                                cell.setCellValue(row.getEbsparaType());             
                                                                }
                                                                if(i==8){
                                                                                cell=       excelRow.createCell(i);
                                                                                cell.setCellValue(row.getParam_name());           
                                                                }
                                                                if(i==9){
                                                                                cell=       excelRow.createCell(i);
                                                                                cell.setCellValue(row.getMinMaxUnitSet());      
                                                                }
                                                                if(i==10){
                                                                                cell=       excelRow.createCell(i);
                                                                                cell.setCellValue(row.getMinValue());   
                                                                }
                                                                if(i==11){
                                                                                cell=       excelRow.createCell(i);
                                                                                cell.setCellValue(row.getMaxValue());  
                                                                }
                                                                if(i==12){
                                                                                cell=       excelRow.createCell(i);
                                                                                cell.setCellValue(row.getActive());          
                                                                }
                                                                if(i==13){
                                                                                cell=       excelRow.createCell(i);
                                                                                cell.setCellValue(row.getMandatory());                
                                                                }
                                                                if(i==14){
                                                                                cell=       excelRow.createCell(i);
                                                                                cell.setCellValue(row.getAllowDropdown());      
                                                                }
                                                                if(i==15){
                                                                                cell=       excelRow.createCell(i);
                                                                                cell.setCellValue(row.getDisplayOn());  
                                                                }
                                                                if(i==16){
                                                                                cell=       excelRow.createCell(i);
                                                                                cell.setCellValue(row.getiOData());         
                                                                }
                                                                if(i==17){
                                                                                cell=       excelRow.createCell(i);
                                                                                cell.setCellValue(row.getTableName());               
                                                                }
                                                                if(i==18){
                                                                                cell=       excelRow.createCell(i);
                                                                                cell.setCellValue(row.getSheet());           
                                                                }
                if(i==0){
                    cell=  excelRow.createCell(i);
                    cell.setCellValue(row.getCaseName());    
             }
             if(i==1){
                    cell=  excelRow.createCell(i);
                    cell.setCellValue(row.getParaName());    
             }
             if(i==2){
                    cell=  excelRow.createCell(i);
                    cell.setCellValue(row.getParaType());    
             }
             if(i==3){
                    cell=  excelRow.createCell(i);
                    cell.setCellValue(row.getParam_label()); 
             }
             if(i==4){
                    cell=  excelRow.createCell(i);
                    cell.setCellValue(row.getUnitSet());     
             }
             if(i==5){
                    cell=  excelRow.createCell(i);
                    cell.setCellValue(row.getParam_type());  
             }
             if(i==6){
                    cell=  excelRow.createCell(i);
                    cell.setCellValue(row.getDisOrder());    
             }
             if(i==7){
                    cell=  excelRow.createCell(i);
                    cell.setCellValue(row.getEbsparaType()); 
             }
             if(i==8){
                    cell=  excelRow.createCell(i);
                    cell.setCellValue(row.getParam_name());  
             }
             if(i==9){
                    cell=  excelRow.createCell(i);
                    cell.setCellValue(row.getMinMaxUnitSet());      
             }
             if(i==10){
                    cell=  excelRow.createCell(i);
                    cell.setCellValue(row.getMinValue());    
             }
             if(i==11){
                    cell=  excelRow.createCell(i);
                    cell.setCellValue(row.getMaxValue());    
             }
              if(i==12){
                    cell=  excelRow.createCell(i);
                    cell.setCellValue(row.getActive());      
             }
             if(i==13){
                    cell=  excelRow.createCell(i);
                    cell.setCellValue(row.getMandatory());   
             }
             if(i==14){
                    cell=  excelRow.createCell(i);
                    cell.setCellValue(row.getAllowDropdown());      
             }
             if(i==15){
                    cell=  excelRow.createCell(i);
                    cell.setCellValue(row.getDisplayOn());   
             }
             if(i==16){
                    cell=  excelRow.createCell(i);
                    cell.setCellValue(row.getiOData());      
             }
             if(i==17){
                    cell=  excelRow.createCell(i);
                    cell.setCellValue(row.getTableName());   
             }
             if(i==18){
                    cell=  excelRow.createCell(i);
                    cell.setCellValue(row.getSheet()); 
             }
                                                                
                                                }
                                }
                                return sheet;
                }
                
                public String downloadProcessPage(DownloadProcess downloadProcess) throws Exception {

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
                    
                    
                  
                                XSSFSheet inputsheet= excelWorkbook.createSheet("Process");
                                
                                try {
                                                int row = 0;
                                                if(downloadProcess != null ){
                                                                Row headerRow = inputsheet.createRow(row);
                                                                headerRow.createCell(0).setCellValue("TEST INFORMATION");
                                                                row = row+1;
                                                                if(downloadProcess.getProcessDtls()!=null){
                                                                                List<ProcessDtls> processDtls = downloadProcess.getProcessDtls();
                                                                                System.out.println("processDtls length :" + processDtls.size());
                                                                                System.out.println(processDtls.get(0).getCustomer());
                                                                                headerRow = inputsheet.createRow(row);
                                                                                headerRow.createCell(0).setCellValue("Model Name");
                                                                                headerRow.createCell(1).setCellValue(processDtls.get(0).getModelName());
                                                                                headerRow.createCell(2).setCellValue("Case Name");
                                                                                headerRow.createCell(3).setCellValue(processDtls.get(0).getCaseName());
                                                                                headerRow.createCell(4).setCellValue("Customers");
                                                                                headerRow.createCell(5).setCellValue(processDtls.get(0).getCustomer());
                                                                                headerRow.createCell(6).setCellValue("Project");
                                                                                headerRow.createCell(7).setCellValue(processDtls.get(0).getProject());
                                                                                row = row +1;
                                                                                headerRow = inputsheet.createRow(row);
                                                                                headerRow.createCell(0).setCellValue("Test Unit");
                                                                                headerRow.createCell(1).setCellValue(processDtls.get(0).getTest_unit());
                                                                                headerRow.createCell(2).setCellValue("Test Date");
                                                                                headerRow.createCell(3).setCellValue(processDtls.get(0).getTest_date());
                                                                                headerRow.createCell(4).setCellValue("Test Time");
                                                                                headerRow.createCell(5).setCellValue(processDtls.get(0).getTest_time());
                                                                                headerRow.createCell(6).setCellValue("Comments");
                                                                                headerRow.createCell(7).setCellValue(processDtls.get(0).getComments());
                                                                                row = row +1;
                                                                                headerRow = inputsheet.createRow(row);
                                                                                headerRow.createCell(0).setCellValue("Test Engineer");
                                                                                headerRow.createCell(1).setCellValue(processDtls.get(0).getTest_engg());
                                                                                headerRow.createCell(2).setCellValue("Analysis Date");
                                                                                headerRow.createCell(3).setCellValue(processDtls.get(0).getAnalysis_date());
                                                                }
                                                                
                                                                if(downloadProcess.getiOData()!=null){
                                                                                
                                                                                List<IOData> inputData = new ArrayList<IOData>();
                                                                                List<IOData> ouputData = new ArrayList<IOData>();
                                                                                
                                                                                for (IOData paramterVO : downloadProcess.getiOData()) {
                                                                                                if(paramterVO.getParam_type().equalsIgnoreCase("I")){
                                                                                                                inputData.add(paramterVO);
                                                                                                }else if(paramterVO.getParam_type().equalsIgnoreCase("O")){
                                                                                                                ouputData.add(paramterVO);
                                                                                                }
                                                                                }
                                                                                row = 6;
                                                                                headerRow = inputsheet.createRow(row);
                                                                                headerRow.createCell(0).setCellValue("GUARANTEE CONDITIONS");
                                                                                row = row+1;
                                                                                headerRow = inputsheet.createRow(row);
                                                                                headerRow.createCell(0).setCellValue("INPUT PARAMETER");
                                                                                headerRow.createCell(1).setCellValue("UNITS");
                                                                                headerRow.createCell(2).setCellValue("TEST CONDITIONS");
                                                                                headerRow.createCell(3).setCellValue("REFERENCE CONDITIONS");
                                                                                row = row+1;
                                                                                for (IOData paramterVO : inputData) {
                                                                                                headerRow = inputsheet.createRow(row);
                                                                                                headerRow.createCell(0).setCellValue(paramterVO.getParam_label());
                                                                                                headerRow.createCell(1).setCellValue(paramterVO.getUnits());
                                                                                                headerRow.createCell(2).setCellValue(paramterVO.getTestCondition());
                                                                                                headerRow.createCell(3).setCellValue(paramterVO.getRefCondition());
                                                                                                row++;
                                                                                }
                                                                                
                                                                                row = row+2;
                                                                                headerRow = inputsheet.createRow(row);
                                                                                headerRow.createCell(0).setCellValue("GUARANTEE DATA");
                                                                                row = row+1;
                                                                                headerRow = inputsheet.createRow(row);
                                                                                headerRow.createCell(0).setCellValue("OUTPUT PARAMETER");
                                                                                headerRow.createCell(1).setCellValue("UNITS");
                                                                                headerRow.createCell(2).setCellValue("TEST CONDITIONS");
                                                                                headerRow.createCell(3).setCellValue("REFERENCE CONDITIONS");
                                                                                row = row+1;
                                                                                for (IOData paramterVO : ouputData) {
                                                                                                headerRow = inputsheet.createRow(row);
                                                                                                headerRow.createCell(0).setCellValue(paramterVO.getParam_label());
                                                                                                headerRow.createCell(1).setCellValue(paramterVO.getUnits());
                                                                                                headerRow.createCell(2).setCellValue(paramterVO.getTestCondition());
                                                                                                headerRow.createCell(3).setCellValue(paramterVO.getRefCondition());
                                                                                                row++;
                                                                                }
                                                                }
                                                                
                                                                row = row+2;
                                                                headerRow = inputsheet.createRow(row);
                                                                headerRow.createCell(0).setCellValue("PERFORMANCE CORRECTIONS");
                                                                row = row+1;
                                                                headerRow = inputsheet.createRow(row);
                                                                headerRow.createCell(0).setCellValue("OUTPUT PARAMETER");
                                                                headerRow.createCell(1).setCellValue("FACTOR");
                                                                headerRow.createCell(2).setCellValue("CORRECTED");
                                                                headerRow.createCell(3).setCellValue("MARGIN");
                                                                row = row+1;
                                                                if(downloadProcess.getParamRefDtls()!=null){
                                                                                List<ParamRefDtls> paramRefDtls = downloadProcess.getParamRefDtls();
                                                                                System.out.println("paramRefDtls : "+ paramRefDtls.size());
                                                                                for (ParamRefDtls paramterVO : downloadProcess.getParamRefDtls()) {
                                                                                                headerRow = inputsheet.createRow(row);
                                                                                                headerRow.createCell(0).setCellValue(paramterVO.getParam_type_value());
                                                                                                headerRow.createCell(1).setCellValue(paramterVO.getParam_factor());
                                                                                                headerRow.createCell(2).setCellValue(paramterVO.getParam_corrected());
                                                                                                headerRow.createCell(3).setCellValue(paramterVO.getParam_margin());
                                                                                                row++;
                                                                                }
                                                                }
                                                }
                                                                
                                                
                                                
                                                
                                                
                                } catch (Exception e) {
                                                // TODO Auto-generated catch block
                                                e.printStackTrace();
                                }
                                
                                //sheet.protectSheet("password");
                                String fileName= "Process.xls";
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
           font.setColor(HSSFColor.BLACK.index);
           font.setFontHeightInPoints((short)15);
           style.setFont(font);
        
           CellStyle hdrStyle = excelWorkbook.createCellStyle();
           hdrStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
       	   hdrStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
           Font hdrFont = excelWorkbook.createFont();
           hdrFont.setColor(HSSFColor.BLACK.index);
           hdrFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
           hdrFont.setFontHeightInPoints((short)12);
           hdrStyle.setFont(hdrFont);
           hdrStyle.setWrapText(true);
        
           CellStyle contentStyle = excelWorkbook.createCellStyle();
           Font contentFont = excelWorkbook.createFont();
           contentFont.setColor(HSSFColor.BLACK.index);
           contentFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
           contentFont.setFontHeightInPoints((short)11);
           contentStyle.setFont(contentFont);
           contentStyle.setWrapText(true);
                  
           XSSFSheet inputsheet= excelWorkbook.createSheet("Data Span");
           try {
                  int row = 0;
                  if(data != null ){
                        Row headerRow = inputsheet.createRow(row);
                        headerRow.createCell(0).setCellValue("DATA SPAN TOOL");
                        headerRow.getCell(0).setCellStyle(style);
                        
                        String date=null;
                        String time=null;
                        if(data.getStartTime()!=null && !data.getStartTime().equalsIgnoreCase("")){
                        	String[] parts = data.getStartTime().split("\\s", 2);
                        	date = parts[0];
                        	time = parts[1]; 
                        }
                        
                        row = row+3;
                        headerRow = inputsheet.createRow(row);
                        headerRow.createCell(0).setCellValue("Model Name");
                        headerRow.getCell(0).setCellStyle(hdrStyle);
                        headerRow.createCell(1).setCellValue(data.getModelName());
                        headerRow.getCell(1).setCellStyle(contentStyle);
                        headerRow.createCell(2).setCellValue("Case Name");
                        headerRow.getCell(2).setCellStyle(hdrStyle);
                        headerRow.createCell(3).setCellValue(data.getCaseName());
                        headerRow.getCell(3).setCellStyle(contentStyle);
                        headerRow.createCell(4).setCellValue("Date of Run");
                        headerRow.getCell(4).setCellStyle(hdrStyle);
                        headerRow.createCell(5).setCellValue(date);
                        headerRow.getCell(5).setCellStyle(contentStyle);
                        headerRow.createCell(6).setCellValue("Time of Run");
                        headerRow.getCell(6).setCellStyle(hdrStyle);
                        headerRow.createCell(7).setCellValue(time); 
                        headerRow.getCell(7).setCellStyle(contentStyle);
                                      
                        if(data.getiOData()!=null){
                               List<ParamDtlsVO> inputData = new ArrayList<ParamDtlsVO>();
                               List<ParamDtlsVO> ouputData = new ArrayList<ParamDtlsVO>();
                                    
                               for (ParamDtlsVO paramterVO : data.getiOData()) {
                                      if(paramterVO.getParam_type().equalsIgnoreCase("I") && paramterVO.getIteration().equalsIgnoreCase("1")){
                                             inputData.add(paramterVO);
                                      }else if(paramterVO.getParam_type().equalsIgnoreCase("O") && paramterVO.getIteration().equalsIgnoreCase("1") && !paramterVO.getParam_label().equalsIgnoreCase("Program Return Status")){
                                             ouputData.add(paramterVO);
                                      }
                               }
                               row = row+2;
                               headerRow = inputsheet.createRow(row);
                               headerRow.createCell(0).setCellValue("Input Parameter(s)");
                               headerRow.getCell(0).setCellStyle(hdrStyle);
                               row = row+1;
                               headerRow = inputsheet.createRow(row);
                               headerRow.createCell(0).setCellValue("Input Parameter");
                               headerRow.createCell(1).setCellValue("Initial Value");
                               headerRow.createCell(2).setCellValue("Final Value");
                               headerRow.createCell(3).setCellValue("Value Increment");
                               for(int i=0;i<=3;i++){
                                      headerRow.getCell(i).setCellStyle(hdrStyle);
                               }
                               row = row+1;
                               for (ParamDtlsVO paramterVO : inputData) {
                                      headerRow = inputsheet.createRow(row);
                                      headerRow.createCell(0).setCellValue(paramterVO.getParam_label());
                                      headerRow.createCell(1).setCellValue(paramterVO.getInitialValue());
                                      headerRow.createCell(2).setCellValue(paramterVO.getFinalValue());
                                      headerRow.createCell(3).setCellValue(paramterVO.getIncreamentValue());
                                      row++;
                               }
                               
                               row = row+2;
                               headerRow = inputsheet.createRow(row);
                               headerRow.createCell(0).setCellValue("Output Parameter");
                               headerRow.getCell(0).setCellStyle(hdrStyle);
                                                              
                               row = row+1;
                               for (ParamDtlsVO paramterVO : ouputData) {
                                      headerRow = inputsheet.createRow(row);
                                      headerRow.createCell(0).setCellValue(paramterVO.getParam_label());
                                      row++;
                               }
                               
                               row = row+2;
                               headerRow = inputsheet.createRow(row);
                               headerRow.createCell(0).setCellValue("Results");
                               headerRow.getCell(0).setCellStyle(hdrStyle);
                               row = row+2;
                               headerRow = inputsheet.createRow(row);
                               headerRow.createCell(0).setCellValue(inputData.get(0).getParam_label()+"(a)");
                               headerRow.getCell(0).setCellStyle(hdrStyle);
                               if(data.getTableViewVO()!=null){
                                      if(data.getTableViewVO().getxValue()!=null){
                                        for(int i=1;i<=data.getTableViewVO().getxValue().size();i++){
                                               String xValue=data.getTableViewVO().getxValue().get(i-1);
                                            headerRow.createCell(i).setCellValue(xValue);
                                            headerRow.getCell(i).setCellStyle(contentStyle);
                                        }
                                      }
                                      row = row+1;
                                      if(data.getTableViewVO().getyValue()!=null){
                                             for (java.util.Map.Entry<String,List<IOValuesVO>> entry : data.getTableViewVO().getyValue().entrySet()) {
                                                    headerRow = inputsheet.createRow(row);
                                                    String key=entry.getKey();
                                                    if(key!=null && !key.equalsIgnoreCase("")){
                                                    headerRow.createCell(0).setCellValue(ouputData.get(0).getParam_label()+"(Y)("+inputData.get(1).getParam_label()+"(b="+key+"))");
                                                    headerRow.getCell(0).setCellStyle(hdrStyle);
                                                    List<IOValuesVO> list=entry.getValue();
                                                       for(int i=1;i<=list.size();i++){
                                                           IOValuesVO iOValuesVO=list.get(i-1);
                                                           headerRow.createCell(i).setCellValue(iOValuesVO.getOutput());
                                                       }
                                                          
                                                     }else{
                                                        headerRow.createCell(0).setCellValue(ouputData.get(0).getParam_label()+"(Y)"); 
                                                        headerRow.getCell(0).setCellStyle(hdrStyle);
                                                        List<IOValuesVO> list=entry.getValue();
                                                           for(int i=1;i<=list.size();i++){
                                                                  IOValuesVO iOValuesVO=list.get(i-1);
                                                                  headerRow.createCell(i).setCellValue(iOValuesVO.getOutput());
                                                           }
                                                     }
                                                    row = row+1;
                                               }
                                        }
                                  }
                               
                               if(data.getCurveFit()!=null){
                            	   if(data.getCurveFit().size()>0){
                            		   for(int i=0;i<data.getCurveFit().size();i++){
                            			   if(i==0){
                            				   row = row+2;
                                               headerRow = inputsheet.createRow(row);
                                               headerRow.createCell(0).setCellValue("Results-Curve Fit");
                                               headerRow.getCell(0).setCellStyle(hdrStyle); 
                                               
                                               row = row+2;
                                               headerRow = inputsheet.createRow(row);
                                               headerRow.createCell(0).setCellValue(inputData.get(0).getParam_label()+"(a)");
                                               headerRow.getCell(0).setCellStyle(hdrStyle);
                                              
                                            	   if(data.getCurveFit().get(i).getxValue()!=null){
                                                       for(int j=1;j<=data.getCurveFit().get(i).getxValue().size();j++){
                                                              String xValue=data.getCurveFit().get(i).getxValue().get(j-1);
                                                           headerRow.createCell(j).setCellValue(xValue);
                                                           headerRow.getCell(j).setCellStyle(contentStyle);
                                                       }
                                                   }
                            			     }
                            			   
                            			   row = row+1;
                                           if(data.getCurveFit().get(i).getyValue()!=null){
                                                  for (java.util.Map.Entry<String,List<IOValuesVO>> entry : data.getCurveFit().get(i).getyValue().entrySet()) {
                                                         headerRow = inputsheet.createRow(row);
                                                         String key=entry.getKey();
                                                         if(key!=null && !key.equalsIgnoreCase("")){
                                                         headerRow.createCell(0).setCellValue(ouputData.get(0).getParam_label()+"(Y)("+inputData.get(1).getParam_label()+"(b="+key+"))");
                                                         headerRow.getCell(0).setCellStyle(hdrStyle);
                                                         List<IOValuesVO> list=entry.getValue();
                                                         for(int j=1;j<=list.size();j++){
                                                                IOValuesVO iOValuesVO=list.get(j-1);
                                                                headerRow.createCell(j).setCellValue(iOValuesVO.getOutput());
                                                         }
                                                               
                                                    }else{
                                                           headerRow.createCell(0).setCellValue(ouputData.get(0).getParam_label()+"(Y)"); 
                                                           headerRow.getCell(0).setCellStyle(hdrStyle);
                                                           List<IOValuesVO> list=entry.getValue();
                                                                for(int j=1;j<=list.size();j++){
                                                                       IOValuesVO iOValuesVO=list.get(j-1);
                                                                       headerRow.createCell(j).setCellValue(iOValuesVO.getOutput());
                                                                }
                                                          }
                                                         row = row+1;
                                                  }
                                           }
                            		   }
                            	   
                            	   }
                            	   row = row+2;
                            	   headerRow = inputsheet.createRow(row);
                            	   headerRow.createCell(0).setCellValue("Curve Fit Equation(s):");
                                   headerRow.getCell(0).setCellStyle(hdrStyle);
                                   
                                   row = row+1;
                            	   for(int i=0;i<data.getCurveFit().size();i++){
                            		    headerRow = inputsheet.createRow(row);
                            		    headerRow.createCell(0).setCellValue(data.getCurveFit().get(i).getEquation());
                                        inputsheet.addMergedRegion(new CellRangeAddress(headerRow.getRowNum(),headerRow.getRowNum(),0,2));
                                        headerRow.createCell(1);
                                        headerRow.createCell(2);
                                          for(int j=0;j<=2;j++){
                                            headerRow.getCell(j).setCellStyle(contentStyle);
                                         }
                                          row = row+1;   
                            	   }
                               }
                         }
                    }
                  
                   for (int i = 0; i < inputsheet.getRow(inputsheet.getLastRowNum()).getPhysicalNumberOfCells(); ++i) {
                        inputsheet.setColumnWidth(i, (int) (inputsheet.getColumnWidth(i) * 3));
                   }
          } catch (Exception e) {
                  e.printStackTrace();
          }
                     
           String fileName= "DataSpan.xls";
           FileOutputStream file = new FileOutputStream(new File(fileName));
           excelWorkbook.write(file);        
           file.close();
           return fileName;
    
    }
                
                
}
