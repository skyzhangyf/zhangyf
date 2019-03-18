package com.tingcream.helloSpringBoot.service;

import java.sql.CallableStatement;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.jdbc.datasource.DataSourceUtils;
import org.springframework.stereotype.Service;

import com.tingcream.helloSpringBoot.Report1Service;
import com.tingcream.helloSpringBoot.util.JdbcUtils;

import oracle.jdbc.OracleTypes;

@Service("report1Service")
public class Report1ServiceImpl implements Report1Service
{
	@Autowired
	private JdbcTemplate jdbcTemplate;

	@Override
	public Map<String, Object> getReport1Data()
	{
		Map<String, Object> txtMap = callReportTxt("pk_report1_txt");
		
		Map<String, Object> reportDatMap = new HashMap<>();
		
		reportDatMap.put("parametersMap", txtMap);
		reportDatMap.put("table1", callProTable("pk_report1_table1"));
		reportDatMap.put("table2", callProTable("pk_report1_table2"));
		reportDatMap.put("table3", callProTable("pk_report1_table3"));
		reportDatMap.put("table4", callProTable("pk_report1_table4"));
		
		return reportDatMap;
	}
	
	private  Map<String, Object> callReportTxt(String proSql) {
		Map<String, Object> resultMap = new HashMap<>();

		Connection con = null;
		CallableStatement cs = null;
		
		try
		{
			con = DataSourceUtils.getConnection(jdbcTemplate.getDataSource());
			
			StringBuilder sbSql = new StringBuilder();
			
			sbSql.append("{call ");
			sbSql.append(proSql);
			sbSql.append("(?)}");

			cs = con.prepareCall(sbSql.toString());
		
			cs.registerOutParameter(1, OracleTypes.CURSOR);// 注册输出参数的类型
			
			cs.execute();
			
			ResultSet rs = (ResultSet) cs.getObject(1);// 获取游标一行的值
			
			while (rs.next())
			{// 转换每行的返回值到Map中
				resultMap.put(rs.getString("KEY"), rs.getString("VALUE"));
			}
			
			JdbcUtils.close(rs);
		}
		catch (SQLException e) {
			resultMap = null;
			
			e.printStackTrace();
		}
		finally
		{
			JdbcUtils.close(cs);
			JdbcUtils.close(con);
		}
		
		// 图片1
		resultMap.put("foreachText1", callProTxtForeach("pk_report1_txt_foreach1"));
		// 图片1
		resultMap.put("picture1", callProPictureLine("pk_report1_picture_line", 1));
		// 图片2
		resultMap.put("picture2", callProPictureLine("pk_report1_picture_line", 2));
		// 图片3
		resultMap.put("picture3", callProPictureLine("pk_report1_picture_line", 3));
		// 图片4
		resultMap.put("picture4", callProPictureLine("pk_report1_picture_line", 4));
		// 图片5
		resultMap.put("picture5", callProPictureLine("pk_report1_picture_line", 5));
		// 图片6
		resultMap.put("picture6", callProPictureLine("pk_report1_picture_line", 6));
		// 图片7
		resultMap.put("picture7", callProPictureLine("pk_report1_picture_line", 7));
				
		resultMap.put("picture8", callProPicturePie("pk_report1_picture8"));
		resultMap.put("picture9", callProPicturePie("pk_report1_picture9"));
		
		return resultMap;
	}
	
	private List<String> callProTxtForeach(String proSql) 
	{
		List<String> resultMapList = new ArrayList<>();
		Connection con = null;
		CallableStatement cs = null;
		
		try
		{
			con = DataSourceUtils.getConnection(jdbcTemplate.getDataSource());
		
			StringBuilder sbSql = new StringBuilder();
			
			sbSql.append("{call ");
			sbSql.append(proSql);
			sbSql.append("(?)}");

			cs = con.prepareCall(sbSql.toString());

			cs.registerOutParameter(1, OracleTypes.CURSOR);// 注册输出参数的类型
			
			cs.execute();
			
			ResultSet rs = (ResultSet) cs.getObject(1);// 获取游标一行的值
			
			while (rs.next())
			{
				resultMapList.add(rs.getString("VALUE"));
			}
			
			JdbcUtils.close(rs);
		}
		catch (SQLException e) {
			resultMapList = null;
		}
		finally
		{
			JdbcUtils.close(cs);
			JdbcUtils.close(con);
		}
		
		return resultMapList;
	}
	
	private List<Map<String, Object>> callProTable(String proSql) 
	{
		List<Map<String, Object>> resultMapList = new ArrayList<>();
		Connection con = null;
		CallableStatement cs = null;
		
		try
		{
			con = DataSourceUtils.getConnection(jdbcTemplate.getDataSource());
		
			StringBuilder sbSql = new StringBuilder();
			
			sbSql.append("{call ");
			sbSql.append(proSql);
			sbSql.append("(?)}");

			cs = con.prepareCall(sbSql.toString());

			cs.registerOutParameter(1, OracleTypes.CURSOR);// 注册输出参数的类型
			
			cs.execute();
			
			ResultSet rs = (ResultSet) cs.getObject(1);// 获取游标一行的值
			
			resultMapList = JdbcUtils.convertResultSetToMapList(rs, true);
			
			JdbcUtils.close(rs);
		}
		catch (SQLException e) {
			resultMapList = null;
		}
		finally
		{
			JdbcUtils.close(cs);
			JdbcUtils.close(con);
		}
		
		return resultMapList;
	}

	private Map<String, Object> callProPictureLine(String proSql, int picNumber)
	{
		Map<String, Object> resultMap = new HashMap<>();
		
		Connection con = null;
		CallableStatement cs = null;
		
		try
		{
			con = DataSourceUtils.getConnection(jdbcTemplate.getDataSource());
			
			StringBuilder sbSql = new StringBuilder();
			
			sbSql.append("{call ");
			sbSql.append(proSql);
			sbSql.append("(?,?,?,?,?)}");

			cs = con.prepareCall(sbSql.toString());

			cs.setInt(1, picNumber);
			cs.registerOutParameter(2, OracleTypes.VARCHAR);
			cs.registerOutParameter(3, OracleTypes.VARCHAR);
			cs.registerOutParameter(4, OracleTypes.VARCHAR);
			cs.registerOutParameter(5, OracleTypes.CURSOR);
			
			cs.execute();
			
			String title = cs.getString(2);
			String valueAxisLabel = cs.getString(3);
			String valueAxisFormat = cs.getString(4);
			
			ResultSet rs = (ResultSet) cs.getObject(5);
			
			List<Map<String, Object>> picMapList = new ArrayList<>();
			
			Map<String, Object> picMap = null;
			
			Map<String, List<Map<String, Object>>> lagndMap = new LinkedHashMap<>();
			
			while (rs.next())
			{
				String legEnd = rs.getString("LEGeND");
				
				List<Map<String, Object>> dataList = null;
				
				if (lagndMap.containsKey(legEnd))
				{
					dataList = lagndMap.get(legEnd);
				}
				else
				{
					dataList = new ArrayList<>();
					
					lagndMap.put(legEnd, dataList);
				}
				
				Map<String, Object> data = new HashMap<>();
				data.put("category", rs.getString("CATEGORY"));
				data.put("value", rs.getDouble("VALUE"));
				
				dataList.add(data);
			}
			
			Iterator<Entry<String, List<Map<String, Object>>>> iterator= lagndMap.entrySet().iterator();  
			  
			while(iterator.hasNext())  
			{  
				Entry<String, List<Map<String, Object>>> entry = iterator.next();
				
				picMap = new HashMap<>();
				
				picMap.put("legend", entry.getKey());
				picMap.put("data", entry.getValue());
				
				picMapList.add(picMap);
			}

			resultMap.put("title", title);
			resultMap.put("valueAxisLabel", valueAxisLabel);
			resultMap.put("valueAxisFormat", valueAxisFormat);
			resultMap.put("data", picMapList);
			resultMap.put("type", "line");
			
			JdbcUtils.close(rs);
		}
		catch (SQLException e) 
		{
			resultMap.put("title", "");
			resultMap.put("valueAxisLabel", "");
			resultMap.put("data", null);
			resultMap.put("type", "line");
		}
		finally
		{
			JdbcUtils.close(con);
			JdbcUtils.close(cs);
		}
		
		return resultMap;
	}
	
	private Map<String, Object> callProPicturePie(String proSql)
	{
		Map<String, Object> resultMap = new HashMap<>();
		
		Connection con = null;
		CallableStatement cs = null;
		
		try
		{
			con = DataSourceUtils.getConnection(jdbcTemplate.getDataSource());
			
			StringBuilder sbSql = new StringBuilder();
			
			sbSql.append("{call ");
			sbSql.append(proSql);
			sbSql.append("(?,?)}");

			cs = con.prepareCall(sbSql.toString());

			cs.registerOutParameter(1, OracleTypes.VARCHAR);// 注册输出参数的类型
			cs.registerOutParameter(2, OracleTypes.CURSOR);// 注册输出参数的类型
			
			cs.execute();
			
			String title = cs.getString(1);
			List<Map<String, Object>> mapList = new ArrayList<>();
			ResultSet rs = (ResultSet) cs.getObject(2);// 获取游标一行的值
			
			while (rs.next())
			{
				Map<String, Object> map = new HashMap<>();
				
				map.put("key", rs.getString("KEY"));
				map.put("value", rs.getString("VALUE"));
				
				mapList.add(map);
			}
			
			resultMap.put("data", mapList);
			resultMap.put("title", title);
			resultMap.put("type", "pie");
			resultMap.put("width", 575);
			resultMap.put("height", 345);
			
			JdbcUtils.close(rs);
		}
		catch (SQLException e) 
		{
			resultMap.put("data", null);
			resultMap.put("title", "");
			resultMap.put("type", "pie");
			resultMap.put("width", 575);
			resultMap.put("height", 345);
		}
		finally
		{
			JdbcUtils.close(con);
			JdbcUtils.close(cs);
		}
		
		return resultMap;
	}
}
