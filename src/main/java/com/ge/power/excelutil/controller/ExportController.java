package com.ge.power.excelutil.controller;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;

import javax.ws.rs.Consumes;
import javax.ws.rs.POST;
import javax.ws.rs.Produces;

import org.apache.commons.codec.binary.Base64;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpEntity;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpMethod;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.client.RestTemplate;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.ge.power.excelutil.dao.WebsilonCCDAO;
import com.ge.power.excelutil.factory.ExportFactory;
import com.ge.power.excelutil.vo.DataSpanVO;
import com.ge.power.excelutil.vo.DownloadProcess;
import com.ge.power.excelutil.vo.ExcelResponse;
import com.ge.power.excelutil.vo.ExcelVO;
import com.ge.power.excelutil.vo.ParamterVO;
import com.google.gson.Gson;

@RestController
@RequestMapping("/services/exportService")

public class ExportController {

	@Autowired
	private ExportFactory exportFactory;
	
	@Autowired
	private WebsilonCCDAO websilonDAO;
	
	//Stage
	private static String url="https://pmeadminstage.run.asv-pr.ice.predix.io/WebsilonAdmin/Services/getIOParams";
	//Prod
	//private static String url="https://pmeadminprod.run.asv-pr.ice.predix.io/WebsilonAdmin/Services/getIOParams";

	@RequestMapping("/dataExport")
	@POST
	@Consumes(MediaType.APPLICATION_JSON_VALUE)
	@Produces(MediaType.APPLICATION_JSON_VALUE)
	public String exportData(@RequestBody List<ParamterVO> paramterVOList){
		String fileName=null;
		try{
			fileName= exportFactory.export(paramterVOList);
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

	@RequestMapping("/downloadExcel")
	@POST
	@Consumes(MediaType.APPLICATION_JSON_VALUE)
	@Produces(MediaType.APPLICATION_JSON_VALUE)
	public String downloadExcel(@RequestBody ExcelVO excelVO){
		String fileName=null;
		System.out.println("downloadExcel :" );
		try{
			excelVO = websilonDAO.getIOParams(excelVO);
			fileName= exportFactory.downloadExcel(excelVO);
			ExcelResponse response = new ExcelResponse();
			response.setOuput(encodeFileToBase64Binary(fileName));
			return new Gson().toJson(response); 

		}catch(Exception e){
			e.printStackTrace();
		}
		return fileName;
	}
	private ExcelVO callRestService(ExcelVO excelVO){
		try{
			ObjectMapper mapper = new ObjectMapper();
			String jsonInput = new Gson().toJson(excelVO);
			System.out.println("jsonInput callRestService  WebsilonAdminProd:" + jsonInput);
	
			RestTemplate restTemplate = new RestTemplate();

			HttpHeaders headers = new HttpHeaders();
			headers.setContentType(MediaType.APPLICATION_JSON);
			HttpEntity<String> entity = new HttpEntity<String>(jsonInput, headers);

			ResponseEntity<String> result = restTemplate.exchange(url, HttpMethod.POST, entity, String.class);
			System.out.println("result.getBody() :" + result.getBody());
			excelVO = mapper.readValue(result.getBody(),ExcelVO.class);
			
		}catch (Exception e) {
			e.printStackTrace();
		}
		return excelVO;		

	}
	
	@RequestMapping("/downloadProcessPage")
	@POST
	@Consumes(MediaType.APPLICATION_JSON_VALUE)
	@Produces(MediaType.APPLICATION_JSON_VALUE)
	public String downloadProcessPage(@RequestBody DownloadProcess downloadProcess){
		String fileName=null;
		System.out.println("downloadProcessPage :" );
		try{
			String jsonInput = new Gson().toJson(downloadProcess);
			System.out.println("jsonInput :" + jsonInput);
			fileName= exportFactory.downloadProcessPage(downloadProcess);
			ExcelResponse response = new ExcelResponse();
			response.setOuput(encodeFileToBase64Binary(fileName));
			return new Gson().toJson(response); 

		}catch(Exception e){
			e.printStackTrace();
		}
		return fileName;
	}
	
	
	@RequestMapping("/downloadDataSpan")
	@POST
	@Consumes(MediaType.APPLICATION_JSON_VALUE)
	@Produces(MediaType.APPLICATION_JSON_VALUE)
	public String downloadDataSpan(@RequestBody DataSpanVO data){
		String fileName=null;
		System.out.println("downloadDataSpan :" );
		try{
			String jsonInput = new Gson().toJson(data);
			System.out.println("jsonInput :" + jsonInput);
			fileName= exportFactory.downloadDataSpan(data);
			ExcelResponse response = new ExcelResponse();
			response.setOuput(encodeFileToBase64Binary(fileName));
			return new Gson().toJson(response); 

		}catch(Exception e){
			e.printStackTrace();
		}
		return fileName;
	}
}
