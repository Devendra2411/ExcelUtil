package com.ge.power.excelutil.util;

public class WebsilonConstant {
	
	public static final String EMPTY_STRING = "";
	public static final String UAA_CHECK_TOKEN_URL = "UAA_CHECK_TOKEN_URL";
	public static final String OAUTH_AUTH_HEADER = "UAA_AUTHORIZATION_HEADER";
	public static final String AUTHORIZATION_HEADER = "Authorization";
	public static final String USERNAME= "user_name";
	public static String validateForNull(String value) {
		if (value == null || "undefined".equals(value) || "null".equals(value)) {
			return "";
		}
		return value;
	}
	
	public static String LEFT_SLASH = "\"";

}
