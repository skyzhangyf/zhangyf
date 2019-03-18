package com.finstone.chart;

import java.awt.Color;
import java.awt.Font;
import java.util.List;
import java.util.Map;

import org.jfree.chart.ChartFactory;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.StandardChartTheme;
import org.jfree.chart.block.BlockBorder;
import org.jfree.chart.labels.StandardPieSectionLabelGenerator;
import org.jfree.chart.plot.PiePlot;
import org.jfree.data.general.DefaultPieDataset;
import org.jfree.ui.RectangleEdge;

public class PieChart
{
	private PieChart()
	{
	}

	public static JFreeChart createPieChart(String title,
			List<Map<String, Object>> dataMapList)
	{
		DefaultPieDataset dataset = new DefaultPieDataset();

		for (Map<String, Object> map : dataMapList)
		{
			String key = map.get("key").toString();
			Double value = Double.valueOf(map.get("value").toString());

			dataset.setValue(key, value);
		}

		JFreeChart chart = createChart(title, dataset);

		return chart;
	}

	public static JFreeChart createChart(String title,
			DefaultPieDataset dataset)
	{
		// 先给个主题，将字体设置好
		Font font = new Font("宋体", Font.PLAIN, 14); // 底部
		StandardChartTheme s = new StandardChartTheme("a");
		s.setRegularFont(font);
		ChartFactory.setChartTheme(s);
		// 2：创建Chart[创建不同图形]
		JFreeChart chart = ChartFactory.createPieChart(title, dataset);
		// 3:设置抗锯齿，防止字体显示不清楚
		ChartUtils.setAntiAlias(chart);// 抗锯齿
		// 4:对柱子进行渲染[[采用不同渲染]]
		ChartUtils.setPieRender(chart.getPlot());//
		PiePlot plot = (PiePlot) chart.getPlot();
		plot.setLabelGenerator(new StandardPieSectionLabelGenerator("{2}"));// 取消图中标签
		plot.setIgnoreZeroValues(true);// 消除数据值为0的项

		plot.setBackgroundAlpha(0);
		plot.setOutlinePaint(Color.WHITE);

		// 设置标注无边框
		chart.getLegend().setFrame(new BlockBorder(Color.WHITE));
		chart.getLegend().setPosition(RectangleEdge.LEFT);

		chart.getLegend().setItemFont(font);
		chart.getTitle().setFont(font);
		// 设置图片边框
		chart.setBorderVisible(true);

		return chart;
	}
}
