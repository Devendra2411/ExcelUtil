package com.ge.power.excelutil.util;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Component;



@Component
public class CommonUtil {
	private static final Logger logger = LoggerFactory.getLogger(CommonUtil.class);
	public static String getSystemEnvVal(String sysValStr){
		try{
			return System.getenv(sysValStr);
		}catch(Exception e){
			logger.error("Exception in getting System Env value for: " + sysValStr + "--Error:: " + e.getMessage());
		}
		return null;
	}
	
}
