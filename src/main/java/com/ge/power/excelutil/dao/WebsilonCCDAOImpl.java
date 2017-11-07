package com.ge.power.excelutil.dao;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Component;

import com.ge.power.excelutil.exception.WebsilonDBTrasactionException;
import com.ge.power.excelutil.util.WebsilonConstant;
import com.ge.power.excelutil.vo.ExcelVO;
import com.ge.power.excelutil.vo.ParamterVO;
import com.ge.power.excelutil.vo.ValidUserVO;

import lombok.extern.slf4j.Slf4j;

@Component
@Slf4j
public class WebsilonCCDAOImpl implements WebsilonCCDAO {

	Logger log = LoggerFactory.getLogger(WebsilonCCDAOImpl.class);
	@Autowired
	protected JdbcTemplate jdbc;

	private static String VAL_USER="select distinct uh.sso,uh.role_id,uh.email_id,r.role_name  from websilon.websn_app_user_hdr uh,websilon.websn_roles r where uh.role_id=r.role_id and uh.active='Y' and uh.sso='";
	private static String GET_PARAMS="select * from websilon.func_get_parameters(";
	
	private static String RIG_COLN = "'";
	
	@SuppressWarnings({ "rawtypes", "unused" })
	public ValidUserVO validUser(ValidUserVO user) throws WebsilonDBTrasactionException {
		log.debug("WebsilonDAOImpl loginValid method :start");
		try{
			String query =VAL_USER+user.getSsoID()+RIG_COLN;
			System.out.println("valid user query :"+ query);
			List<Map<String, Object>> results = jdbc.queryForList(query);
			if(results != null && results.size()>0){
				for (Map row1 : results) {
					user.setValidUser("Yes");
				}
			}else{
				user.setValidUser("No");
			}

		}catch(Exception e){
			log.warn("Exception at loginValid () :"+e);
			throw new WebsilonDBTrasactionException("WEB-EX001","WebsilonDBException",e.getMessage());
		}

		return user;
	}
	
	@SuppressWarnings("rawtypes")
	@Override
	public ExcelVO getIOParams(ExcelVO excelVO) throws WebsilonDBTrasactionException {
		ExcelVO uploadFile=new ExcelVO();
		List<ParamterVO> paramterList = new ArrayList<ParamterVO>();
		ParamterVO param=null;
		   try{         	  
			String query=GET_PARAMS+excelVO.getModelID()+","+excelVO.getCaseID()+")";	
			System.out.println("Query::"+query);
			List<Map<String, Object>> list = jdbc.queryForList(query);
			for (Map row1 : list) {
				param= new ParamterVO();
				param.setCaseName(WebsilonConstant.validateForNull((String)row1.get("case_name")).trim());
				param.setParaName(WebsilonConstant.validateForNull((String)row1.get("param_value_type")).trim());
				param.setParaType(WebsilonConstant.validateForNull((String)row1.get("param_data_Type")).trim());
				param.setParam_label(WebsilonConstant.validateForNull((String)row1.get("ui_name")).trim());
				param.setUnitSet(WebsilonConstant.validateForNull((String)row1.get("unit_set")).trim());
				param.setParam_type(WebsilonConstant.validateForNull((String)row1.get("input_output_flag")).trim());
				if(row1.get("display_order") != null && !(new Integer((int)row1.get("display_order")).toString()).equalsIgnoreCase("") ){
					param.setDisOrder(new Integer((int)row1.get("display_order")).toString());
				}else{
					param.setDisOrder("");
				}
				param.setEbsparaType(WebsilonConstant.validateForNull((String)row1.get("io_name")).trim());
				param.setParam_name(WebsilonConstant.validateForNull((String)row1.get("ebsn_sys_name")).trim());
				param.setMinMaxUnitSet(WebsilonConstant.validateForNull((String)row1.get("")).trim());
				param.setMinValue(WebsilonConstant.validateForNull((String)row1.get("min_value")).trim());
				param.setMaxValue(WebsilonConstant.validateForNull((String)row1.get("max_value")).trim());
				param.setActive(WebsilonConstant.validateForNull((String)row1.get("active")).trim());
				param.setMandatory(WebsilonConstant.validateForNull((String)row1.get("mandatory_input")).trim());
				param.setAllowDropdown(WebsilonConstant.validateForNull((String)row1.get("formula_1")).trim());
				param.setDisplayOn(WebsilonConstant.validateForNull((String)row1.get("displayon")).trim());
				param.setiOData(WebsilonConstant.validateForNull((String)row1.get("category_name")).trim());
				param.setTableName(WebsilonConstant.validateForNull((String)row1.get("param_type")).trim());
				param.setSheet(WebsilonConstant.validateForNull((String)row1.get("")).trim());
				param.setParam_value(WebsilonConstant.validateForNull((String)row1.get("value")).trim());
				param.setParam_unit(WebsilonConstant.validateForNull((String)row1.get("")).trim());
				paramterList.add(param);						  
			}
					
		}catch(Exception e){
			log.warn("Exception at getIOParams () :"+e);
			e.printStackTrace();
			throw new WebsilonDBTrasactionException("WEB-EX001","WebsilonDBException",e.getMessage());
		}
		   uploadFile.setiOData(paramterList);	
		return uploadFile;
	}

}
