package com.finstone.word;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;

public class WordUtils
{
	private WordUtils () {}
	
	public static void createReport(String templatePath, String outPath, Map<String, Object> wordDataMap) {
		FileInputStream fis = null;
		FileOutputStream fos = null;
		
		try
		{
			// 读取word模板，必须是docx类型
			File file = new File(templatePath);
	
			fis = new FileInputStream(file);
			WordTemplate template = new WordTemplate(fis);
	
			// 替换数据
			template.replaceDocument(wordDataMap);
			
			// 生成文件
			File outputFile = new File(outPath);
			
			fos = new FileOutputStream(outputFile);
			
			template.getDocument().write(fos);
		}
		catch (IOException e)
		{
			System.out.println("生成word报表失败");
			
			e.printStackTrace();
		}
		catch (Exception e)
		{
			System.out.println("生成word报表失败");
			
			e.printStackTrace();
		}
		finally
		{
			if (fos != null) {
				try
				{
					fos.close();
				} 
				catch (IOException e)
				{
					e.printStackTrace();
				}
			}

			if (fis != null) {
				try
				{
					fis.close();
				} 
				catch (IOException e)
				{
					e.printStackTrace();
				}
			}
		}
	}
}
