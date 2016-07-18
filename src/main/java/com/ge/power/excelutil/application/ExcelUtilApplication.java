package com.ge.power.excelutil.application;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.EnableAutoConfiguration;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.context.web.SpringBootServletInitializer;
import org.springframework.context.annotation.ComponentScan;





@SpringBootApplication
@ComponentScan({"com.ge.power.excelutil"})
@EnableAutoConfiguration
public class ExcelUtilApplication extends SpringBootServletInitializer {

	
    public static void main(String[] args) {
    	Object[] obj = {
    			ExcelUtilApplication.class,
    			};
    SpringApplication.run(obj,args);
    
    
  
    }
   
}


