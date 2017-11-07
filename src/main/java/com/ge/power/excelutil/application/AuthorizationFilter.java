package com.ge.power.excelutil.application;

import java.io.IOException;

import javax.servlet.Filter;
import javax.servlet.FilterChain;
import javax.servlet.FilterConfig;
import javax.servlet.ServletException;
import javax.servlet.ServletRequest;
import javax.servlet.ServletResponse;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpEntity;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.stereotype.Component;
import org.springframework.util.LinkedMultiValueMap;
import org.springframework.util.MultiValueMap;
import org.springframework.web.client.HttpClientErrorException;
import org.springframework.web.client.RestTemplate;

import com.ge.power.excelutil.dao.WebsilonCCDAO;
import com.ge.power.excelutil.exception.InvalidTokenException;
import com.ge.power.excelutil.exception.UnauthorizedException;
import com.ge.power.excelutil.exception.WebsilonDBTrasactionException;
import com.ge.power.excelutil.util.WebsilonConstant;
import com.ge.power.excelutil.vo.ValidUserVO;


/**
 * Filter to authorize requests
 *
 */
@Component
class AuthorizationFilter implements Filter {

	private static final Logger logger = LoggerFactory.getLogger(AuthorizationFilter.class);

	private final String checkTokenURL = "https://a99b7fee-a495-4161-89c3-faa83054627d.predix-uaa.run.asv-pr.ice.predix.io/check_token";//CommonUtil.getSystemEnvVal(WebsilonConstant.UAA_CHECK_TOKEN_URL);
	private String authorizationHeader = "Basic cG93ZXJfcG1lX3N0YWdlOnptM2kxVTBKbFRXV2hCaGY=";//CommonUtil.getSystemEnvVal(WebsilonConstant.OAUTH_AUTH_HEADER);

	@Autowired
	private WebsilonCCDAO websilonDAO;
	
	@Override
	public void init(FilterConfig arg0) throws ServletException {

		// Default implementation ignored
	}

	@Override
	public void doFilter(ServletRequest req, ServletResponse res, FilterChain chain)
			throws IOException, ServletException {
		HttpServletRequest httpServletRequest = (HttpServletRequest) req;
		String accessToken = httpServletRequest.getHeader(WebsilonConstant.AUTHORIZATION_HEADER);
		System.out.println("-accessToken-" + accessToken );
		//String path = httpServletRequest.getRequestURI().substring(httpServletRequest.getContextPath().length());
		//if(!path.contains("/health")){
			if (accessToken != null) {
				if (accessToken.contains("Bearer ")) {
					accessToken = accessToken.replace("Bearer ", "");
				}
				RestTemplate restTemplate = new RestTemplate();
				MultiValueMap<String, String> body = new LinkedMultiValueMap<>();
				body.add("token", accessToken);
				HttpHeaders headers = new HttpHeaders();
				headers.add(HttpHeaders.CONTENT_TYPE, MediaType.APPLICATION_FORM_URLENCODED_VALUE);
				headers.add(HttpHeaders.AUTHORIZATION, authorizationHeader);
				
				System.out.println("-checkTokenURL-" + checkTokenURL );
				System.out.println("-authorizationHeader-" + authorizationHeader );
				
				HttpEntity<?> httpEntity = new HttpEntity<>(body, headers);
				try {
					String out = restTemplate.postForObject(checkTokenURL, httpEntity, String.class);
					JSONParser parser = new JSONParser();
					JSONObject jsonObject;
					jsonObject = (JSONObject) parser.parse(out);
					String sso = (String) jsonObject.get(WebsilonConstant.USERNAME);
					System.out.println("- sso -" + sso);
					if (sso != null && !sso.trim().isEmpty()) {
						//String userSSO = httpServletRequest.getHeader("userSSO");
						//if (userSSO != null && userSSO.equalsIgnoreCase(sso)) {
						ValidUserVO access = new ValidUserVO();
						access.setSsoID(sso.trim());
						try {
							access=websilonDAO.validUser(access);
						} catch (WebsilonDBTrasactionException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						}
						if(access.getValidUser().equalsIgnoreCase("Yes")){
							HttpServletResponse response = (HttpServletResponse) res;
							response.setHeader("Access-Control-Allow-Origin", "*"); //$NON-NLS-1$ //$NON-NLS-2$
							response.setHeader("Access-Control-Allow-Methods", "POST, GET, OPTIONS, DELETE ,PUT"); //$NON-NLS-1$//$NON-NLS-2$
							response.setHeader("Access-Control-Max-Age", "3600"); //$NON-NLS-1$ //$NON-NLS-2$
							response.setHeader("Access-Control-Allow-Headers",
									"Origin,Accept,X-Requested-With,Content-Type,Access-Control-Request-Method,Access-Control-Request-Headers,Authorization,userSSO");
							chain.doFilter(new XSSRequestWrapper(httpServletRequest), res);
						}else{
							logger.info("SSO in header and SSO in check token response does not match");
							throw new UnauthorizedException("Unauthorized due to SSO tampered in Header");
						}
						
							
						/*} else {
							logger.info("SSO in header and SSO in check token response does not match");
							throw new UnauthorizedException("Unauthorized due to SSO tampered in Header");
						}*/
					} else {
						throw new InvalidTokenException("Token Expired");
					}
				} catch (org.json.simple.parser.ParseException e) {
					logger.info("Unable to Parse token");
					throw new InvalidTokenException("Invalid token : Unable to parse");
				} catch (HttpClientErrorException e) {
					throw new InvalidTokenException("Invalid token");
				}
			} else {
				throw new UnauthorizedException("Unauthorized Access");
			}
		//}
	}

	@Override
	public void destroy() {

		// Default implementation ignored
	}
	
}