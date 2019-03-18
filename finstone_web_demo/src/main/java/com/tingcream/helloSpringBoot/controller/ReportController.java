package com.tingcream.helloSpringBoot.controller;

import java.util.Map;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;

import com.finstone.word.WordUtils;
import com.tingcream.helloSpringBoot.Report1Service;

@Controller
public class ReportController
{
	@Autowired
	private Report1Service report1Service;

	// 读取word模板，必须是docx类型
	private static final String TEMPLATE_PATH = "/apply/reportTemplate1.docx";
	// 输出报表word文件路径
	private static final String OUT_PATH = "/apply/输出1.docx";
	
	@ResponseBody
	@RequestMapping("/report1.html")
	public String index()
	{
		System.out.println("报表生成开始！");

		// 存储报表全部数据
		Map<String, Object> wordDataMap = report1Service.getReport1Data();

		 WordUtils.createReport(TEMPLATE_PATH, OUT_PATH, wordDataMap);

		System.out.println("报表生成结束！");

		return "hello SpringBoot 牛逼!!";
	}
}