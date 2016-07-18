package com.ge.power.excelutil.controller;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;

import javax.servlet.http.HttpServletResponse;
import javax.ws.rs.Consumes;
import javax.ws.rs.POST;
import javax.ws.rs.Produces;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.MediaType;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

import com.ge.power.excelutil.factory.ExportFactory;
import com.ge.power.excelutil.vo.ExcelResponse;
import com.ge.power.excelutil.vo.ParamterVO;
import com.google.gson.Gson;

import java.io.File;

import org.apache.commons.codec.binary.Base64;

@RestController
@RequestMapping("/services/exportService")

public class ExportController {

@Autowired
private ExportFactory exportFactory;


@RequestMapping("/dataExport")
@POST
@Consumes(MediaType.APPLICATION_JSON_VALUE)
@Produces(MediaType.APPLICATION_JSON_VALUE)
public String exportData(@RequestBody List<ParamterVO> paramterVOList){
	String fileName=null;
   // InputStream is = null;
    //response.setHeader("Content-disposition", "attachment;filenameFinanceDashBoardd.xls");
  //  response.setContentType("application/vnd.ms-excel");
	
	try{
		 fileName= exportFactory.export(paramterVOList);
		//  is=new FileInputStream(fileName);
	   //   org.apache.commons.io.IOUtils.copy(is, response.getOutputStream());
	   //   response.flushBuffer();
		 ExcelResponse response = new ExcelResponse();
		 response.setOuput(encodeFileToBase64Binary(fileName));
		 return new Gson().toJson(response); 
		
	}catch(Exception e){
		e.printStackTrace();
	}
	return fileName;
	}

@RequestMapping(value = "/getGreeting", method = RequestMethod.GET)
public String getGreeting(@RequestParam(value="name") String name) {
	String result="Hello "+name;		
	return result;
}	

@SuppressWarnings("unused")
private String encodeFileToBase64Binary(String fileName)
		throws IOException {

	File file = new File(fileName);
	byte[] bytes = loadFile(file);
	byte[] encoded = Base64.encodeBase64(bytes);
	String encodedString = new String(encoded);

	return encodedString;
}
@SuppressWarnings("resource")
private static byte[] loadFile(File file) throws IOException {
    InputStream is = new FileInputStream(file);

    long length = file.length();
    if (length > Integer.MAX_VALUE) {
        // File is too large
    }
    byte[] bytes = new byte[(int)length];
    
    int offset = 0;
    int numRead = 0;
    while (offset < bytes.length
           && (numRead=is.read(bytes, offset, bytes.length-offset)) >= 0) {
        offset += numRead;
    }

    if (offset < bytes.length) {
        throw new IOException("Could not completely read file "+file.getName());
    }

    is.close();
    return bytes;
}
}
