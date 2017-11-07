package com.ge.power.excelutil.dao;

import com.ge.power.excelutil.exception.WebsilonDBTrasactionException;
import com.ge.power.excelutil.vo.ExcelVO;
import com.ge.power.excelutil.vo.ValidUserVO;

public interface WebsilonCCDAO {
	public ValidUserVO validUser(ValidUserVO user) throws WebsilonDBTrasactionException;
	public ExcelVO getIOParams(ExcelVO excelVO)throws WebsilonDBTrasactionException;
}
