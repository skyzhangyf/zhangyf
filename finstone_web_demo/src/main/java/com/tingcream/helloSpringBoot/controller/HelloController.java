package com.tingcream.helloSpringBoot.controller;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;

import com.finstone.word.WordUtils;

@Controller
public class HelloController
{
	@ResponseBody
	@RequestMapping("/index.html")
	public String index()
	{

		System.out.println("报表生成开始！");

		// 读取word模板，必须是docx类型
		 String templatePath = "D:\\word\\reportTemplate1.docx";
//		String templatePath = "/apply/reportTemplate1.docx";
		// 输出报表word文件路径
		 String outPath = "D:\\word\\输出1.docx";
//		String outPath = "/apply/输出1.docx";
		// 存储报表全部数据
		Map<String, Object> wordDataMap = getWordDataMap();

		 WordUtils.createReport(templatePath, outPath, wordDataMap);

		System.out.println("报表生成结束！");

		return "hello SpringBoot 牛逼!!";
	}
	
	private  Map<String, Object> getWordDataMap()
	{
		// 存储报表中不循环的数据
		Map<String, Object> parametersMap = setParametersMap();
		// 存储报表中循环表格1
		List<Map<String, Object>> table1 = setTable1();
		// 存储报表中循环表格2
		List<Map<String, Object>> table2 = setTable2();
		// 存储报表中循环表格3
		List<Map<String, Object>> table3 = setTable3();
		// 存储报表中循环表格4
		List<Map<String, Object>> table4 = setTable4();

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
		parametersMap.put("picture8", setPicture8());
		// 图片9
		parametersMap.put("picture9", setPicture9());

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

	private  Map<String, Object> setPicture8()
	{
		String title = "单位性质";

		List<Map<String, Object>> picMapList = new ArrayList<>();

		Map<String, Object> picMap1 = new HashMap<>();

		picMap1.put("key", "国家机关、事业单位");
		picMap1.put("value", 79.13);

		picMapList.add(picMap1);

		Map<String, Object> picMap2 = new HashMap<>();

		picMap2.put("key", "国有企业");
		picMap2.put("value", 15.86);

		picMapList.add(picMap2);

		Map<String, Object> picMap3 = new HashMap<>();

		picMap3.put("key", "城镇私营企业");
		picMap3.put("value", 1.63);

		picMapList.add(picMap3);

		Map<String, Object> picMap4 = new HashMap<>();

		picMap4.put("key", "城镇集体企业");
		picMap4.put("value", 2.5);

		picMapList.add(picMap4);

		Map<String, Object> picMap5 = new HashMap<>();

		picMap5.put("key", "外商投资企业");
		picMap5.put("value", 0.5);

		picMapList.add(picMap5);

		Map<String, Object> picture8 = new HashMap<>();

		picture8.put("title", title);
		picture8.put("data", picMapList);
		picture8.put("type", "pie");

		return picture8;
	}

	private  Map<String, Object> setPicture9()
	{
		String title = "单位经济类型";

		List<Map<String, Object>> picMapList = new ArrayList<>();

		Map<String, Object> picMap1 = new HashMap<>();

		picMap1.put("key", "其他私有");
		picMap1.put("value", 0);
		picMapList.add(picMap1);

		Map<String, Object> picMap2 = new HashMap<>();
		picMap2.put("key", "其他联营");
		picMap2.put("value", 1);
		picMapList.add(picMap2);

		Map<String, Object> picMap3 = new HashMap<>();
		picMap3.put("key", "内资");
		picMap3.put("value", 0);
		picMapList.add(picMap3);

		Map<String, Object> picMap4 = new HashMap<>();
		picMap4.put("key", "国外投资");
		picMap4.put("value", 2);
		picMapList.add(picMap4);

		Map<String, Object> picMap5 = new HashMap<>();
		picMap5.put("key", "国有与集体联营");
		picMap5.put("value", 7.51);
		picMapList.add(picMap5);

		Map<String, Object> picMap6 = new HashMap<>();
		picMap6.put("key", "国有全资");
		picMap5.put("value", 15.52);
		picMapList.add(picMap5);

		Map<String, Object> picMap7 = new HashMap<>();
		picMap7.put("key", "国有独资(公司)");
		picMap7.put("value", 0);
		picMapList.add(picMap7);

		Map<String, Object> picMap8 = new HashMap<>();
		picMap8.put("key", "国有联营");
		picMap8.put("value", 3);
		picMapList.add(picMap8);

		Map<String, Object> picMap9 = new HashMap<>();
		picMap9.put("key", "外资");
		picMap9.put("value", 0);
		picMapList.add(picMap9);

		Map<String, Object> picMap10 = new HashMap<>();
		picMap10.put("key", "有限责任(公司)");
		picMap10.put("value", 0);
		picMapList.add(picMap10);

		Map<String, Object> picMap11 = new HashMap<>();
		picMap11.put("key", "港、澳、台投资");
		picMap11.put("value", 0);
		picMapList.add(picMap11);

		Map<String, Object> picMap12 = new HashMap<>();
		picMap12.put("key", "港、澳、台投资股份有限(公司)");
		picMap12.put("value", 46.86);
		picMapList.add(picMap12);

		Map<String, Object> picMap13 = new HashMap<>();
		picMap13.put("key", "私有");
		picMap13.put("value", 4);
		picMapList.add(picMap13);

		Map<String, Object> picMap14 = new HashMap<>();
		picMap14.put("key", "私有合伙");
		picMap14.put("value", 12.48);
		picMapList.add(picMap14);

		Map<String, Object> picMap = new HashMap<>();
		picMap.put("key", "私有独资");
		picMap.put("value", 0);
		picMapList.add(picMap);

		Map<String, Object> picMap15 = new HashMap<>();
		picMap15.put("key", "私营有限责任(公司)");
		picMap15.put("value", 0);
		picMapList.add(picMap15);

		Map<String, Object> picMap16 = new HashMap<>();
		picMap16.put("key", "私有股份有限(公司)");
		picMap16.put("value", 0);
		picMapList.add(picMap16);

		Map<String, Object> picMap17 = new HashMap<>();
		picMap17.put("key", "联营");
		picMap17.put("value", 0);
		picMapList.add(picMap17);

		Map<String, Object> picMap18 = new HashMap<>();
		picMap18.put("key", "股份合作");
		picMap18.put("value", 0);
		picMapList.add(picMap18);

		Map<String, Object> picMap19 = new HashMap<>();
		picMap19.put("key", "股份有限(公司)");
		picMap19.put("value", 9.16);
		picMapList.add(picMap19);

		Map<String, Object> picMap20 = new HashMap<>();
		picMap20.put("key", "集体合资");
		picMap20.put("value", 2.0);
		picMapList.add(picMap20);

		Map<String, Object> picMap21 = new HashMap<>();
		picMap21.put("key", "城镇集体企业");
		picMap21.put("value", 2.5);
		picMapList.add(picMap21);

		Map<String, Object> picMap22 = new HashMap<>();
		picMap22.put("key", "外商投资企业");
		picMap22.put("value", 0.5);
		picMapList.add(picMap22);

		Map<String, Object> picture9 = new HashMap<>();

		picture9.put("title", title);
		picture9.put("data", picMapList);
		picture9.put("type", "pie");
		picture9.put("width", 575);
		picture9.put("height", 345);

		return picture9;
	}

	private  Map<String, Object> setParametersMap()
	{
		Map<String, Object> parametersMap = new HashMap<>();
		
		// 需要进行文本替换的信息
		parametersMap.put("text1", "V02.81");
		parametersMap.put("text2", "-100.0%");
		parametersMap.put("text3", "9.22");
		parametersMap.put("text4", "-99.65%");
		parametersMap.put("text5", "19.19");
		parametersMap.put("text6", "-99.76%");
		parametersMap.put("text7", "16.9");
		parametersMap.put("text8", "0.0%");
		parametersMap.put("text9", "0.0");
		parametersMap.put("text10", "4,339");
		parametersMap.put("text11", "116.09");
		parametersMap.put("text12", "37.32");
		parametersMap.put("text13", "116.09");
		parametersMap.put("text14", "30.93");
		parametersMap.put("text15", "6.38");
		parametersMap.put("text16", "0.0");
		parametersMap.put("text17", "4.9633");
		parametersMap.put("text18", "128.11");
		parametersMap.put("text19", "90.65");
		parametersMap.put("text20", "116");
		parametersMap.put("text21", "0");
		parametersMap.put("text22", "0");
		parametersMap.put("text23", "0.0");
		parametersMap.put("text24", "417,981.71");
		parametersMap.put("text25", "29");
		parametersMap.put("text26", "0");
		parametersMap.put("text27", "2");
		parametersMap.put("text28", "394");
		parametersMap.put("text29", "0");
		parametersMap.put("text30", "100.00%");
		parametersMap.put("text31", "35,464");
		parametersMap.put("text32", "373,155");
		parametersMap.put("text33", "0.27");
		parametersMap.put("text34", "7");
		parametersMap.put("text35", "79");
		parametersMap.put("text36", "0");
		parametersMap.put("text37", "78");
		parametersMap.put("text38", "13,800.0");
		parametersMap.put("text39", "94");
		parametersMap.put("text40", "1,380.0");
		parametersMap.put("text41", "423");
		parametersMap.put("text42", "4,453");
		parametersMap.put("text43", "0");
		parametersMap.put("text44", "0");
		parametersMap.put("text45", "2,187");
		parametersMap.put("text46", "【】【杜新田】【000029549】【02】【2000-01-01】");
		parametersMap.put("text47", "【】【许玉清】【000029551】【02】【2000-01-01】");
		parametersMap.put("text48", "【】【蒋广军】【000029554】【02】【2000-01-01】");
		parametersMap.put("text49", "【】【王金立】【000029555】【02】【2000-01-01】");
		parametersMap.put("text50", "【】【许伟峰】【000029557】【02】【2000-01-01】");
		parametersMap.put("table1_text1", "2018年12月111");

		List<String> l = new ArrayList<>();

		l.add("【2429.40】【10560.00】【022924771】【2016-04-25】");
		l.add("【1760.16】【12000.00】【023035076】【2016-05-13】");
		l.add("【2157.84】【12000.00】【023150415】【2016-05-27】");

		parametersMap.put("foreachText1", l);

		return parametersMap;
	}

	private  List<Map<String, Object>> setTable1()
	{
		List<Map<String, Object>> table1 = new ArrayList<>();

		Map<String, Object> map11 = new HashMap<>();
		map11.put("bh", "101");
		map11.put("km", "公积金存款");
		map11.put("qc", "1137351258.49");
		map11.put("qm", "1058799665.62");

		table1.add(map11);

		Map<String, Object> map12 = new HashMap<>();
		map12.put("bh", "101");
		map12.put("km", "公积金存款");
		map12.put("qc", "930595589.70");
		map12.put("qm", "953132528.97");

		table1.add(map12);

		Map<String, Object> map13 = new HashMap<>();
		map13.put("bh", "101");
		map13.put("km", "公积金存款");
		map13.put("qc", "941851079.10");
		map13.put("qm", "930595589.70");

		table1.add(map13);

		Map<String, Object> map14 = new HashMap<>();
		map14.put("bh", "101");
		map14.put("km", "公积金存款");
		map14.put("qc", "1058799665.62");
		map14.put("qm", "1069434824.68");

		table1.add(map14);

		return table1;
	}

	private  List<Map<String, Object>> setTable2()
	{
		List<Map<String, Object>> table2 = new ArrayList<>();

		Map<String, Object> map21 = new HashMap<>();
		map21.put("bh", "101");
		map21.put("km", "公积金存款");
		map21.put("qc", "1069434824.68");
		map21.put("qm", "1000636205.83");
		table2.add(map21);

		Map<String, Object> map22 = new HashMap<>();
		map22.put("bh", "101");
		map22.put("km", "公积金存款");
		map22.put("qc", "1000636205.83");
		map22.put("qm", "941851097.10");
		table2.add(map22);

		Map<String, Object> map23 = new HashMap<>();
		map23.put("bh", "102");
		map23.put("km", "增值收益存款");
		map23.put("qc", "162676662.00");
		map23.put("qm", "162783240.24");
		table2.add(map23);

		Map<String, Object> map24 = new HashMap<>();
		map24.put("bh", "102");
		map24.put("km", "增值收益存款");
		map24.put("qc", "162676662.00");
		map24.put("qm", "162676662.00");
		table2.add(map24);

		Map<String, Object> map25 = new HashMap<>();
		map25.put("bh", "102");
		map25.put("km", "增值收益存款");
		map25.put("qc", "125.60091.27");
		map25.put("qm", "125.60091.27");
		table2.add(map25);

		Map<String, Object> map26 = new HashMap<>();
		map26.put("bh", "102");
		map26.put("km", "增值收益存款");
		map26.put("qc", "162676662.00");
		map26.put("qm", "162676662.00");
		table2.add(map26);

		Map<String, Object> map27 = new HashMap<>();
		map27.put("bh", "102");
		map27.put("km", "增值收益存款");
		map27.put("qc", "125.60091.27");
		map27.put("qm", "125.60091.27");
		table2.add(map27);

		Map<String, Object> map28 = new HashMap<>();
		map28.put("bh", "102");
		map28.put("km", "增值收益存款");
		map28.put("qc", "125.60091.27");
		map28.put("qm", "162676662.00");
		table2.add(map28);

		Map<String, Object> map29 = new HashMap<>();
		map29.put("bh", "111");
		map29.put("km", "应收利息");
		map29.put("qc", "122998464.00");
		map29.put("qm", "122998464.00");
		table2.add(map29);

		Map<String, Object> map210 = new HashMap<>();
		map210.put("bh", "111");
		map210.put("km", "应收利息");
		map210.put("qc", "122998464.00");
		map210.put("qm", "122998464.00");
		table2.add(map210);

		Map<String, Object> map211 = new HashMap<>();
		map211.put("bh", "111");
		map211.put("km", "应收利息");
		map211.put("qc", "122998464.00");
		map211.put("qm", "122998464.00");
		table2.add(map211);

		Map<String, Object> map212 = new HashMap<>();
		map212.put("bh", "111");
		map212.put("km", "应收利息");
		map212.put("qc", "122998464.00");
		map212.put("qm", "122998464.00");
		table2.add(map212);

		return table2;
	}

	private  List<Map<String, Object>> setTable3()
	{
		List<Map<String, Object>> table3 = new ArrayList<>();

		Map<String, Object> map31 = new HashMap<>();
		map31.put("bh", "A");
		map31.put("hy", "农、林、牧、渔业");
		map31.put("dws", "178");
		map31.put("rs", "6748");
		map31.put("ye", "20,749.86");
		table3.add(map31);

		Map<String, Object> map32 = new HashMap<>();

		map32.put("bh", "B");
		map32.put("hy", " 采矿业");
		map32.put("dws", " 2");
		map32.put("rs", " 4");
		map32.put("ye", " 2.41");
		table3.add(map32);

		Map<String, Object> map33 = new HashMap<>();
		map33.put("bh", "C");
		map33.put("hy", " 制造业");
		map33.put("dws", " 1150");
		map33.put("rs", " 87047");
		map33.put("ye", " 216,510.18");
		table3.add(map33);

		Map<String, Object> map34 = new HashMap<>();
		map34.put("bh", "D");
		map34.put("hy", " 电力、燃气及水的生产和供应");
		map34.put("dws", " 77");
		map34.put("rs", " 8355");
		map34.put("ye", " 49,928.92");
		table3.add(map34);

		Map<String, Object> map35 = new HashMap<>();
		map35.put("bh", "E");
		map35.put("hy", " 建筑业");
		map35.put("dws", " 138");
		map35.put("rs", " 2412");
		map35.put("ye", " 5,460.14");
		table3.add(map35);

		Map<String, Object> map36 = new HashMap<>();
		map36.put("bh", "F");
		map36.put("hy", " 交通运输、仓储和邮政");
		map36.put("dws", " 196");
		map36.put("rs", " 11638");
		map36.put("ye", " 50,188.71");
		table3.add(map36);

		Map<String, Object> map37 = new HashMap<>();
		map37.put("bh", "G");
		map37.put("hy", " 信息传输、计算机服务和软件业");
		map37.put("dws", " 58");
		map37.put("rs", " 3737");
		map37.put("ye", " 23,798.77");
		table3.add(map37);

		Map<String, Object> map38 = new HashMap<>();
		map38.put("bh", "H");
		map38.put("hy", " 批发和零售业");
		map38.put("dws", " 52");
		map38.put("rs", " 1282");
		map38.put("ye", " 3,498.77");
		table3.add(map38);

		Map<String, Object> map39 = new HashMap<>();
		map39.put("bh", "I");
		map39.put("hy", " 住宿和餐饮业");
		map39.put("dws", " 16");
		map39.put("rs", " 164");
		map39.put("ye", " 214.52");
		table3.add(map39);

		Map<String, Object> map40 = new HashMap<>();
		map40.put("bh", "J");
		map40.put("hy", " 金融业");
		map40.put("dws", " 208");
		map40.put("rs", " 13228");
		map40.put("ye", " 105,161.77");
		table3.add(map40);

		Map<String, Object> map41 = new HashMap<>();
		map41.put("bh", "K");
		map41.put("hy", " 房地产业");
		map41.put("dws", " 235");
		map41.put("rs", " 3300");
		map41.put("ye", " 5,748.10");
		table3.add(map41);

		Map<String, Object> map42 = new HashMap<>();
		map42.put("bh", "L");
		map42.put("hy", " 租赁和商务服务");
		map42.put("dws", " 58");
		map42.put("rs", " 1076");
		map42.put("ye", " 1,189.36");
		table3.add(map42);

		Map<String, Object> map43 = new HashMap<>();
		map43.put("bh", "M");
		map43.put("hy", " 科学研究、技术服务和地质勘查");
		map43.put("dws", " 36");
		map43.put("rs", " 526");
		map43.put("ye", " 2,158.01");
		table3.add(map43);

		Map<String, Object> map44 = new HashMap<>();
		map44.put("bh", "N");
		map44.put("hy", " 水利、环境和公共设施管理");
		map44.put("dws", " 189");
		map44.put("rs", " 7675");
		map44.put("ye", " 26,873.92");
		table3.add(map44);

		Map<String, Object> map45 = new HashMap<>();
		map45.put("bh", "O");
		map45.put("hy", " 居民服务和其他服务");
		map45.put("dws", " 121");
		map45.put("rs", " 6441");
		map45.put("ye", " 22,265.55");
		table3.add(map45);

		Map<String, Object> map46 = new HashMap<>();
		map46.put("bh", "P");
		map46.put("hy", " 教育");
		map46.put("dws", " 536");
		map46.put("rs", " 57566");
		map46.put("ye", " 231,703.29");
		table3.add(map46);

		Map<String, Object> map47 = new HashMap<>();
		map47.put("bh", "Q");
		map47.put("hy", " 卫生、社会保障和社会福利");
		map47.put("dws", " 380");
		map47.put("rs", " 38125");
		map47.put("ye", " 125,239.51");
		table3.add(map47);

		Map<String, Object> map48 = new HashMap<>();
		map48.put("bh", "R");
		map48.put("hy", " 文化、体验和");
		map48.put("dws", " 123");
		map48.put("rs", " 3495");
		map48.put("ye", " 18,154.87");
		table3.add(map48);

		Map<String, Object> map49 = new HashMap<>();

		map49.put("bh", "S");
		map49.put("hy", " 公共管理、社会保障和社会组织");
		map49.put("dws", " 1134");
		map49.put("rs", " 56487");
		map49.put("ye", " 252,048.50");
		table3.add(map49);

		return table3;
	}

	private  List<Map<String, Object>> setTable4()
	{
		List<Map<String, Object>> table4 = new ArrayList<>();

		Map<String, Object> map411 = new HashMap<>();
		map411.put("km", "101");
		map411.put("kmmc", "住房公积金存款");
		map411.put("kmjb", "1");
		map411.put("kmsx", "01");
		map411.put("kmyefx", "02");
		table4.add(map411);

		Map<String, Object> map412 = new HashMap<>();
		map412.put("km", "10101");
		map412.put("kmmc", "活期");
		map412.put("kmjb", "2");
		map412.put("kmsx", "01");
		map412.put("kmyefx", "02");
		table4.add(map412);

		Map<String, Object> map413 = new HashMap<>();
		map413.put("km", "10102");
		map413.put("kmmc", "定期款");
		map413.put("kmjb", "2");
		map413.put("kmsx", "01");
		map413.put("kmyefx", "02");
		table4.add(map413);

		Map<String, Object> map414 = new HashMap<>();
		map414.put("km", "102");
		map414.put("kmmc", "增值收益存款");
		map414.put("kmjb", "1");
		map414.put("kmsx", "01");
		map414.put("kmyefx", "02");
		table4.add(map414);

		Map<String, Object> map415 = new HashMap<>();
		map415.put("km", "10201");
		map415.put("kmmc", "活期");
		map415.put("kmjb", "2");
		map415.put("kmsx", "01");
		map415.put("kmyefx", "02");
		table4.add(map415);

		Map<String, Object> map416 = new HashMap<>();
		map416.put("km", "10202");
		map416.put("kmmc", "定期款");
		map416.put("kmjb", "2");
		map416.put("kmsx", "01");
		map416.put("kmyefx", "02");
		table4.add(map416);

		return table4;
	}
}