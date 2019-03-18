package com.finstone.chart;

import java.awt.Color;
import java.text.DecimalFormat;
import java.util.List;
import java.util.Map;

import org.jfree.chart.ChartFactory;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.axis.NumberAxis;
import org.jfree.chart.axis.ValueAxis;
import org.jfree.chart.block.BlockBorder;
import org.jfree.data.category.DefaultCategoryDataset;

/**
 * 
 * 折线图 创建图表步骤：<br/>
 * 1：创建数据集合<br/>
 * 2：创建Chart：<br/>
 * 3:设置抗锯齿，防止字体显示不清楚<br/>
 * 4:对柱子进行渲染，<br/>
 * 5:对其他部分进行渲染<br/>
 * 6:使用chartPanel接收<br/>
 * 
 */
public class LineChart
{
	private LineChart()
	{
	}

	public static JFreeChart createLineChart(String title,
			String valueAxisLabel, List<Map<String, Object>> dataMapList)
	{
		// 1：创建数据集合
		DefaultCategoryDataset dataset = new DefaultCategoryDataset();

		for (Map<String, Object> dataMap : dataMapList)
		{
			String legend = dataMap.get("legend").toString();

			List<Map<String, Object>> dataList = (List<Map<String, Object>>) dataMap.get("data");

			for (Map<String, Object> map : dataList)
			{
				String category = map.get("category").toString();
				String value = map.get("value").toString();
				if (ChartUtils.isPercent(value))
				{
					value = value.substring(0, value.length() - 1);
				}
				if (ChartUtils.isNumber(value))
				{
					dataset.setValue(Double.parseDouble(value), legend, category);
				}
			}
		}

		JFreeChart chart = createChart(title, valueAxisLabel, dataset);

		return chart;
	}

	private static JFreeChart createChart(String title, String valueAxisLabel,
			DefaultCategoryDataset dataset)
	{
		// 2：创建Chart[创建不同图形]
		JFreeChart chart = ChartFactory.createLineChart(title, "",
				valueAxisLabel, dataset);
		// 3:设置抗锯齿，防止字体显示不清楚
		ChartUtils.setAntiAlias(chart);// 抗锯齿
		// 4:对柱子进行渲染[[采用不同渲染]]
		ChartUtils.setLineRender(chart.getCategoryPlot(), false, true);//
		// 5:对其他部分进行渲染
		ChartUtils.setXAixs(chart.getCategoryPlot());// X坐标轴渲染
		ChartUtils.setYAixs(chart.getCategoryPlot());// Y坐标轴渲染
		// 设置标注无边框
		chart.getLegend().setFrame(new BlockBorder(Color.WHITE));

		NumberAxis numberAxis = (NumberAxis) chart.getCategoryPlot()
				.getRangeAxis();
		numberAxis.setNumberFormatOverride(new DecimalFormat("0"));

		// 设置图片边框
		chart.setBorderVisible(true);
		
		ValueAxis rAxis = chart.getCategoryPlot().getRangeAxis();
		rAxis.setLabelAngle(0);//标题内容显示角度（1.6时候水平） 
		rAxis.setLabelPaint(Color.red);//标题内容颜色 
		return chart;
	}
}
