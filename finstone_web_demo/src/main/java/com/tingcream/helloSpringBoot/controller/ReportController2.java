package com.tingcream.helloSpringBoot.controller;

import java.sql.CallableStatement;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.jdbc.datasource.DataSourceUtils;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;

import com.finstone.word.WordUtils;
import com.tingcream.helloSpringBoot.Report1Service;

import oracle.jdbc.OracleTypes;

@Controller
public class ReportController2
{

	@Autowired
	private JdbcTemplate jdbcTemplate;
	
	@Autowired
	private Report1Service report1Service;

	@ResponseBody
	@RequestMapping("/report.html")
	public String index()
	{

		System.out.println("报表生成开始！");

		// 读取word模板，必须是docx类型
//		 String templatePath = "D:\\word\\reportTemplate1.docx";
		String templatePath = "/apply/reportTemplate1.docx";
		// 输出报表word文件路径
//		 String outPath = "D:\\word\\输出1.docx";
		String outPath = "/apply/输出1.docx";
		// 存储报表全部数据
		Map<String, Object> wordDataMap = getWordDataMap();

		 WordUtils.createReport(templatePath, outPath, wordDataMap);

		System.out.println("报表生成结束！");

		return "hello SpringBoot 牛逼!!";
	}

	private  Map<String, Object> callReport1Txt() {
		Map<String, Object> resultMap = new HashMap<>();

		Connection con = null;
		CallableStatement cs = null;
		
		try
		{
			con = DataSourceUtils.getConnection(jdbcTemplate.getDataSource());
			
			String sql = "{call pk_report1_txt(?)}";// 调用的sql
			cs = con.prepareCall(sql);
		
			cs.registerOutParameter(1, OracleTypes.CURSOR);// 注册输出参数的类型
			
			cs.execute();
			
			ResultSet rs = (ResultSet) cs.getObject(1);// 获取游标一行的值
			
			while (rs.next())
			{// 转换每行的返回值到Map中
				resultMap.put(rs.getString("KEY"), rs.getString("VALUE"));
			}
			
			rs.close();
		}
		catch (SQLException e) {
			e.printStackTrace();
		}
		catch (Exception e) {
			e.printStackTrace();
		}
		finally
		{
			if (cs != null) 
			{
				try
				{
					cs.close();
				} catch (SQLException e)
				{
					cs = null;
				}
			}
			
			if (con != null)
			{
				try
				{
					con.close();
				} catch (SQLException e)
				{
					con = null;
				}
			}
		}
		
		return resultMap;
	}
	
	private  List<Map<String, Object>> callProTable1() {
		List<Map<String, Object>> mapList = new ArrayList<>();
		Connection con = null;
		CallableStatement cs = null;
		
		try
		{
			con = DataSourceUtils.getConnection(jdbcTemplate.getDataSource());
		
			String sql = "{call pk_report1_table1(?)}";// 调用的sql
			cs = con.prepareCall(sql);

			cs.registerOutParameter(1, OracleTypes.CURSOR);// 注册输出参数的类型
			
			cs.execute();
			
			ResultSet rs = (ResultSet) cs.getObject(1);// 获取游标一行的值
			
			while (rs.next())
			{
				Map<String, Object> map = new HashMap<>();
				
				map.put("bh", rs.getString("BH"));
				map.put("km", rs.getString("KM"));
				map.put("qc", rs.getString("QC"));
				map.put("qm", rs.getString("QM"));
				
				mapList.add(map);
			}
			
			rs.close();
		}
		catch (SQLException e) {
			e.printStackTrace();
		}
		catch (Exception e) {
			e.printStackTrace();
		}
		finally
		{
			if (cs != null) 
			{
				try
				{
					cs.close();
				} catch (SQLException e)
				{
					cs = null;
				}
			}
			
			if (con != null)
			{
				try
				{
					con.close();
				} catch (SQLException e)
				{
					con = null;
				}
			}
		}
		
		return mapList;
	}
	
	private  List<Map<String, Object>> callProTable2() {
		List<Map<String, Object>> mapList = new ArrayList<>();
		Connection con = null;
		CallableStatement cs = null;
		
		try
		{
			con = DataSourceUtils.getConnection(jdbcTemplate.getDataSource());
		
			String sql = "{call pk_report1_table2(?)}";// 调用的sql
			cs = con.prepareCall(sql);

			cs.registerOutParameter(1, OracleTypes.CURSOR);// 注册输出参数的类型
			
			cs.execute();
			
			ResultSet rs = (ResultSet) cs.getObject(1);// 获取游标一行的值
			
			while (rs.next())
			{
				Map<String, Object> map = new HashMap<>();
				
				map.put("bh", rs.getString("BH"));
				map.put("km", rs.getString("KM"));
				map.put("qc", rs.getString("QC"));
				map.put("qm", rs.getString("QM"));
				
				mapList.add(map);
			}
			
			rs.close();
		}
		catch (SQLException e) {
			e.printStackTrace();
		}
		catch (Exception e) {
			e.printStackTrace();
		}
		finally
		{
			if (cs != null) 
			{
				try
				{
					cs.close();
				} catch (SQLException e)
				{
					cs = null;
				}
			}
			
			if (con != null)
			{
				try
				{
					con.close();
				} catch (SQLException e)
				{
					con = null;
				}
			}
		}
		
		return mapList;
	}
	
	private  List<Map<String, Object>> callProTable3() {
		List<Map<String, Object>> mapList = new ArrayList<>();
		Connection con = null;
		CallableStatement cs = null;
		
		try
		{
			con = DataSourceUtils.getConnection(jdbcTemplate.getDataSource());
		
			String sql = "{call pk_report1_table3(?)}";// 调用的sql
			cs = con.prepareCall(sql);

			cs.registerOutParameter(1, OracleTypes.CURSOR);// 注册输出参数的类型
			
			cs.execute();
			
			ResultSet rs = (ResultSet) cs.getObject(1);// 获取游标一行的值
			
			while (rs.next())
			{
				Map<String, Object> map = new HashMap<>();
				
				map.put("bh", rs.getString("BH"));
				map.put("hy", rs.getString("HY"));
				map.put("dws", rs.getString("DWS"));
				map.put("rs", rs.getString("RS"));
				map.put("ye", rs.getString("YE"));
				
				mapList.add(map);
			}
			
			rs.close();
		}
		catch (SQLException e) {
			e.printStackTrace();
		}
		catch (Exception e) {
			e.printStackTrace();
		}
		finally
		{
			if (cs != null) 
			{
				try
				{
					cs.close();
				} catch (SQLException e)
				{
					cs = null;
				}
			}
			
			if (con != null)
			{
				try
				{
					con.close();
				} catch (SQLException e)
				{
					con = null;
				}
			}
		}
		
		return mapList;
	}
	
	private  List<Map<String, Object>> callProTable4() {
		List<Map<String, Object>> mapList = new ArrayList<>();
		Connection con = null;
		CallableStatement cs = null;
		
		try
		{
			con = DataSourceUtils.getConnection(jdbcTemplate.getDataSource());
		
			String sql = "{call pk_report1_table4(?)}";// 调用的sql
			cs = con.prepareCall(sql);

			cs.registerOutParameter(1, OracleTypes.CURSOR);// 注册输出参数的类型
			
			cs.execute();
			
			ResultSet rs = (ResultSet) cs.getObject(1);// 获取游标一行的值
			
			while (rs.next())
			{
				Map<String, Object> map = new HashMap<>();
				
				map.put("km", rs.getString("KM"));
				map.put("kmmc", rs.getString("KMMC"));
				map.put("kmjb", rs.getString("KMJB"));
				map.put("kmsx", rs.getString("KMSX"));
				map.put("kmyefx", rs.getString("KMYEFX"));
				
				mapList.add(map);
			}
			
			rs.close();
		}
		catch (SQLException e) {
			e.printStackTrace();
		}
		catch (Exception e) {
			e.printStackTrace();
		}
		finally
		{
			if (cs != null) 
			{
				try
				{
					cs.close();
				} catch (SQLException e)
				{
					cs = null;
				}
			}
			
			if (con != null)
			{
				try
				{
					con.close();
				} catch (SQLException e)
				{
					con = null;
				}
			}
		}
		
		return mapList;
	}
	
	private  Map<String, Object> callProPicture8() {
		Map<String, Object> resultMap = new HashMap<>();
		
		Connection con = null;
		CallableStatement cs = null;
		
		try
		{
			con = DataSourceUtils.getConnection(jdbcTemplate.getDataSource());
		
			String sql = "{call pk_report1_picture8(?,?)}";// 调用的sql
			cs = con.prepareCall(sql);

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
			
			rs.close();
		}
		catch (SQLException e) {
			e.printStackTrace();
		}
		catch (Exception e) {
			e.printStackTrace();
		}
		finally
		{
			if (cs != null) 
			{
				try
				{
					cs.close();
				} catch (SQLException e)
				{
					cs = null;
				}
			}
			
			if (con != null)
			{
				try
				{
					con.close();
				} catch (SQLException e)
				{
					con = null;
				}
			}
		}
		
		return resultMap;
	}
	
	private  Map<String, Object> callProPicture9() {
		Map<String, Object> resultMap = new HashMap<>();
		
		Connection con = null;
		CallableStatement cs = null;
		
		try
		{
			con = DataSourceUtils.getConnection(jdbcTemplate.getDataSource());
		
			String sql = "{call pk_report1_picture9(?,?)}";// 调用的sql
			cs = con.prepareCall(sql);

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
			
			rs.close();
		}
		catch (SQLException e) {
			e.printStackTrace();
		}
		catch (Exception e) {
			e.printStackTrace();
		}
		finally
		{
			if (cs != null) 
			{
				try
				{
					cs.close();
				} catch (SQLException e)
				{
					cs = null;
				}
			}
			
			if (con != null)
			{
				try
				{
					con.close();
				} catch (SQLException e)
				{
					con = null;
				}
			}
		}
		
		return resultMap;
	}
	
	private Map<String, Object> executePro() throws SQLException{
		Map<String, Object> resultMap = new HashMap<>();

		Connection con = DataSourceUtils.getConnection(jdbcTemplate.getDataSource());
		
		String sql = "{call pk_report1(?,?,?)}";// 调用的sql
		CallableStatement cs = con.prepareCall(sql);
		
		try
		{
			cs.registerOutParameter(1, OracleTypes.INTEGER);// 设置输出参数的类型
			cs.registerOutParameter(2, OracleTypes.VARCHAR);// 设置输出参数的类型
			cs.registerOutParameter(3, OracleTypes.CURSOR);// 注册输出参数的类型
			
			cs.execute();
			
			ResultSet rs = (ResultSet) cs.getObject(3);// 获取游标一行的值
			
			List list = new ArrayList();
			while (rs.next())
			{// 转换每行的返回值到Map中
				Map rowMap = new HashMap();
				rowMap.put("A", rs.getString("A"));
				rowMap.put("B", rs.getString("B"));
				
				list.add(rowMap);
			}
			
			resultMap.put("table1", list);
			
			rs.close();
		}
		finally
		{
			cs.close();
			con.close();
		}
		
		return resultMap;
	}
	
	private  Map<String, Object> getWordDataMap()
	{
		// 存储报表中不循环的数据
		Map<String, Object> parametersMap = setParametersMap();
		// 存储报表中循环表格1
		List<Map<String, Object>> table1 = callProTable1();
		// 存储报表中循环表格2
		List<Map<String, Object>> table2 = callProTable2();
		// 存储报表中循环表格3
		List<Map<String, Object>> table3 = callProTable3();
		// 存储报表中循环表格4
		List<Map<String, Object>> table4 = callProTable4();

		// 图片1
		parametersMap.put("picture1", setPicture1());
		// 图片2
		parametersMap.put("picture2", setPicture2());
		// 图片3
		parametersMap.put("picture3", setPicture3());
		// 图片4
		parametersMap.put("picture4", setPicture4());
		// 图片5
		parametersMap.put("picture5", setPicture5());
		// 图片6
		parametersMap.put("picture6", setPicture6());
		// 图片7
		parametersMap.put("picture7", setPicture7());
		// 图片8
		parametersMap.put("picture8", callProPicture8());
		// 图片9
		parametersMap.put("picture9", callProPicture9());

		// 存储报表全部数据
		Map<String, Object> wordDataMap = new HashMap<>();

		wordDataMap.put("parametersMap", parametersMap);
		wordDataMap.put("table1", table1);
		wordDataMap.put("table2", table2);
		wordDataMap.put("table3", table3);
		wordDataMap.put("table4", table4);

		return wordDataMap;
	}

	private  Map<String, Object> setPicture1()
	{
		String title = "归集业务情况";
		String valueAxisLabel = "归集额（万元）";
		List<String> categories = new ArrayList<>();

		categories.add("2017(Q4)");
		categories.add("2018(Q4)");

		List<Map<String, Object>> picMapList = new ArrayList<>();

		Map<String, Object> picMap1 = new HashMap<>();

		picMap1.put("legend", "季度");

		List<Map<String, Object>> dataList11 = new ArrayList<>();

		Map<String, Object> data111 = new HashMap<>();
		data111.put("category", "2017(Q4)");
		data111.put("value", 290000D);
		dataList11.add(data111);

		Map<String, Object> data112 = new HashMap<>();
		data112.put("category", "2018(Q4)");
		data112.put("value", 190000D);
		dataList11.add(data112);

		Map<String, Object> data113 = new HashMap<>();
		data113.put("category", "2019(Q4)");
		data113.put("value", 170000D);
		dataList11.add(data113);

		picMap1.put("data", dataList11);

		picMapList.add(picMap1);

		Map<String, Object> picMap2 = new HashMap<>();

		picMap2.put("legend", "年度");

		List<Map<String, Object>> dataList21 = new ArrayList<>();

		Map<String, Object> data211 = new HashMap<>();
		data211.put("category", "2017(Q4)");
		data211.put("value", 170000D);
		dataList21.add(data211);

		Map<String, Object> data212 = new HashMap<>();
		data212.put("category", "2018(Q4)");
		data212.put("value", 0D);
		dataList21.add(data212);

		Map<String, Object> data213 = new HashMap<>();
		data213.put("category", "2019(Q4)");
		data213.put("value", 90000D);
		dataList21.add(data213);

		picMap2.put("data", dataList21);

		picMapList.add(picMap2);

		Map<String, Object> picture1 = new HashMap<>();

		picture1.put("title", title);
		picture1.put("valueAxisLabel", valueAxisLabel);
		picture1.put("data", picMapList);
		picture1.put("type", "line");

		return picture1;
	}

	private  Map<String, Object> setPicture2()
	{
		String title = "提取业务情况";
		String valueAxisLabel = "提取额（万元）";
		List<String> categories = new ArrayList<>();

		categories.add("2017(Q4)");
		categories.add("2018(Q4)");

		List<Map<String, Object>> picMapList = new ArrayList<>();

		Map<String, Object> picMap1 = new HashMap<>();

		picMap1.put("legend", "季度");

		List<Map<String, Object>> dataList11 = new ArrayList<>();

		Map<String, Object> data111 = new HashMap<>();
		data111.put("category", "2017(Q4)");
		data111.put("value", 160000D);
		dataList11.add(data111);

		Map<String, Object> data112 = new HashMap<>();
		data112.put("category", "2018(Q4)");
		data112.put("value", 110000D);
		dataList11.add(data112);

		picMap1.put("data", dataList11);

		picMapList.add(picMap1);

		Map<String, Object> picture1 = new HashMap<>();

		picture1.put("title", title);
		picture1.put("valueAxisLabel", valueAxisLabel);
		picture1.put("data", picMapList);
		picture1.put("type", "line");

		return picture1;
	}

	private  Map<String, Object> setPicture3()
	{
		String title = "贷款业务情况";
		String valueAxisLabel = "贷款额（万元）";
		List<String> categories = new ArrayList<>();

		categories.add("2017(Q4)");
		categories.add("2018(Q4)");

		List<Map<String, Object>> picMapList = new ArrayList<>();

		Map<String, Object> picMap1 = new HashMap<>();

		picMap1.put("legend", "季度");

		List<Map<String, Object>> dataList11 = new ArrayList<>();

		Map<String, Object> data111 = new HashMap<>();
		data111.put("category", "2017(Q4)");
		data111.put("value", 300000D);
		dataList11.add(data111);

		Map<String, Object> data112 = new HashMap<>();
		data112.put("category", "2018(Q4)");
		data112.put("value", 150000D);
		dataList11.add(data112);

		picMap1.put("data", dataList11);

		picMapList.add(picMap1);

		Map<String, Object> picture1 = new HashMap<>();

		picture1.put("title", title);
		picture1.put("valueAxisLabel", valueAxisLabel);
		picture1.put("data", picMapList);
		picture1.put("type", "line");

		return picture1;
	}

	private  Map<String, Object> setPicture4()
	{
		String title = "保值增值情况";
		String valueAxisLabel = "保值增值额（万元）";
		List<String> categories = new ArrayList<>();

		categories.add("2017(Q4)");
		categories.add("2018(Q4)");

		List<Map<String, Object>> picMapList = new ArrayList<>();

		Map<String, Object> picMap1 = new HashMap<>();

		picMap1.put("legend", "季度");

		List<Map<String, Object>> dataList11 = new ArrayList<>();

		Map<String, Object> data112 = new HashMap<>();
		data112.put("category", "2018(Q4)");
		data112.put("value", 1300D);
		dataList11.add(data112);

		picMap1.put("data", dataList11);

		picMapList.add(picMap1);

		Map<String, Object> picture1 = new HashMap<>();

		picture1.put("title", title);
		picture1.put("valueAxisLabel", valueAxisLabel);
		picture1.put("data", picMapList);
		picture1.put("type", "line");

		return picture1;
	}

	private  Map<String, Object> setPicture5()
	{
		String title = "资产风险情况";
		String valueAxisLabel = "逾期率(%)";

		List<Map<String, Object>> picMapList = new ArrayList<>();

		Map<String, Object> picMap1 = new HashMap<>();

		picMap1.put("legend", "季度");

		List<Map<String, Object>> dataList11 = new ArrayList<>();

		Map<String, Object> data112 = new HashMap<>();
		data112.put("category", "2018(Q4)");
		data112.put("value", 0.89D);
		dataList11.add(data112);

		picMap1.put("data", dataList11);

		picMapList.add(picMap1);

		Map<String, Object> picture1 = new HashMap<>();

		picture1.put("title", title);
		picture1.put("valueAxisLabel", valueAxisLabel);
		picture1.put("valueAxisFormat", "#0.0%");
		picture1.put("data", picMapList);
		picture1.put("type", "line");

		return picture1;
	}

	private  Map<String, Object> setPicture6()
	{
		String title = "其他应收款情况";
		String valueAxisLabel = "其他应收款(万元)";

		List<Map<String, Object>> picMapList = new ArrayList<>();

		Map<String, Object> picMap1 = new HashMap<>();

		picMap1.put("legend", "季度");

		List<Map<String, Object>> dataList11 = new ArrayList<>();

		Map<String, Object> data112 = new HashMap<>();
		data112.put("category", "2018(Q4)");
		data112.put("value", 0D);
		dataList11.add(data112);

		picMap1.put("data", dataList11);

		picMapList.add(picMap1);

		Map<String, Object> picture1 = new HashMap<>();

		picture1.put("title", title);
		picture1.put("valueAxisLabel", valueAxisLabel);
		picture1.put("data", picMapList);
		picture1.put("type", "line");

		return picture1;
	}

	private  Map<String, Object> setPicture7()
	{
		String title = "其他应付款情况";
		String valueAxisLabel = "其他应付款(万元)";

		List<Map<String, Object>> picMapList = new ArrayList<>();

		Map<String, Object> picMap1 = new HashMap<>();

		picMap1.put("legend", "季度");

		List<Map<String, Object>> dataList11 = new ArrayList<>();

		Map<String, Object> data112 = new HashMap<>();
		data112.put("category", "2018(Q4)");
		data112.put("value", 0D);
		dataList11.add(data112);

		picMap1.put("data", dataList11);

		picMapList.add(picMap1);

		Map<String, Object> picture1 = new HashMap<>();

		picture1.put("title", title);
		picture1.put("valueAxisLabel", valueAxisLabel);
		picture1.put("data", picMapList);
		picture1.put("type", "line");

		return picture1;
	}

	private  Map<String, Object> setParametersMap()
	{
		Map<String, Object> parametersMap = callReport1Txt();
			
		List<String> l = new ArrayList<>();

		l.add("【2429.40】【10560.00】【022924771】【2016-04-25】");
		l.add("【1760.16】【12000.00】【023035076】【2016-05-13】");
		l.add("【2157.84】【12000.00】【023150415】【2016-05-27】");

		parametersMap.put("foreachText1", l);

		return parametersMap;
	}
}