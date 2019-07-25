package msword.report;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.hibernate.transform.Transformers;

import msword.Report;
import msword.report.table.Section;
import msword.report.table.section.TemplateRow;

public class Operate {
	
	
	
	/**
	 * 解析
	 * @param report
	 */
	public static void execute(Report report) {
		
		if (report != null && report.document != null)
		{			
			textParse(report, null, null, null, null, null, report.varValues);
			
			msword.report.root.Operate.execute(report);
			if (report.smartBaseDao != null 
					&& report.rootList != null 
					&& report.rootList.size() > 0)
			{
				for (int i=0; i<report.rootList.size(); i++)
				{
					Root root = report.rootList.get(i);
					if (root != null 
							&& root.sql != null 
							&& !root.sql.trim().equals("") 
							&& root.placeholder != null 
							&& !root.placeholder.trim().equals(""))
					{
						List<Map> list = report.smartBaseDao.getSqlQuery(root.sql).setMaxResults(1).setResultTransformer(Transformers.ALIAS_TO_ENTITY_MAP).list();
						if (list != null && list.size() > 0 && list.get(0) != null)
						{
							Map varValues = list.get(0);
							if (report.varValues == null)
							{
								report.varValues = new HashMap<String, Object>();
							}
//							report.varValues.put(root.placeholder, varValues);
							root.varValues = varValues;
							
							textParse(report, root.placeholder, root.dateFormat, root.datetimeFormat, root.fieldsOfDatetime, root.isomer, varValues);
						}
						
					}
				}
				
				
			}	
			//删除root数据标签
			msword.report.root.Operate.removeTag(report);
			
			
			
			List<Table> tables = getTables(report);			
			msword.report.Operate.reload(report);
			
			if (tables != null && tables.size() > 0)
			{
				for (int i=0; i<tables.size(); i++)
				{
					Table table = tables.get(i);
					//解析数据
					msword.report.table.Operate.analyticData(table);
					//根据配置合并数据行
					msword.report.table.Operate.merge(table);
					//删除数据区域开始结束标签
					msword.report.table.Operate.removeTag(table);
					
				}
			
			}
			
			
		}
		
	}
	
	/**
	 * 解析输入文本
	 * @param report
	 * @param placeholder  数据占位符
	 * @param dateFormat 日期格式字符串 需要符合java.text.SampleDateFormat规范
	 * @param datetimeFormat 日期时间格式字符串 需要符合java.text.SimpleDateFormat规范
	 * @param fieldsOfDatetime 日期时间字段名称序列,多个已英文逗号分开， 形如    deptname,username 
	 * @param isomer 异构样式配置
	 * @param varValues 存放变量占位符及对应值的map
	 * @return
	 */
	public static void textParse(Report report, String placeholder, String dateFormat, String datetimeFormat, String fieldsOfDatetime, com.alibaba.fastjson.JSONArray isomer, Map<String, Object> varValues) {
		if (report != null && report.document != null)
		{			
			List<XWPFParagraph> paragraphList = report.document.getParagraphs();
			if (paragraphList != null && paragraphList.size() > 0)
			{
				for (int i3=0;i3<paragraphList.size(); i3++)
				{
					XWPFParagraph paragraph = paragraphList.get(i3);
					if (paragraph != null)
					{
						List<XWPFRun> runs = paragraph.getRuns();	
						
						msword.Operate.pictureAnaly(varValues, placeholder, runs);

						String result = msword.Operate.runsParse(placeholder, dateFormat, datetimeFormat, fieldsOfDatetime, isomer, paragraph.getText(), runs, varValues);
						
					}
				}
			}
			
			List<XWPFTable> tableList = report.document.getTables();
			if (tableList != null && tableList.size() > 0)
			{
				for (int i=0; i<tableList.size(); i++)
				{
					XWPFTable xwpfTable = tableList.get(i);
					if (xwpfTable != null && xwpfTable.getRows() != null && xwpfTable.getRows().size() > 0)
					{
						for (int i2=0; i2<xwpfTable.getRows().size(); i2++)
						{
							XWPFTableRow row = xwpfTable.getRows().get(i2);
							if (row != null && row.getTableCells() != null && row.getTableCells().size() > 0)
							{
								for (int i4=0; i4<row.getTableCells().size(); i4++)
								{
									paragraphList = row.getCell(i4).getParagraphs();
									if (paragraphList != null && paragraphList.size() > 0)
									{
										for (int i5=0;i5<paragraphList.size(); i5++)
										{
											XWPFParagraph paragraph = paragraphList.get(i5);
											if (paragraph != null)
											{
												List<XWPFRun> runs = paragraph.getRuns();	
												
												msword.Operate.pictureAnaly(varValues, placeholder, runs);
	
												String result = msword.Operate.runsParse(placeholder, dateFormat, datetimeFormat, fieldsOfDatetime, isomer, paragraph.getText(), runs, varValues);
												
											}
										}
									
									}
									
								}
								
							}
						}
						
					}
				}
			}
			
		

		}

	}
	
	/**
	 * 再替换没有匹配的${}为空字符
	 * @param report
	 */
	public static void scavengeSuperfluous(Report report) {
		
		if (report != null && report.document != null)
		{
			List<XWPFParagraph> paragraphList = report.document.getParagraphs();
			if (paragraphList != null && paragraphList.size() > 0)
			{
				for (int i3=0;i3<paragraphList.size(); i3++)
				{
					XWPFParagraph paragraph = paragraphList.get(i3);
					if (paragraph != null)
					{
						List<XWPFRun> runs = paragraph.getRuns();	

						String result = msword.Operate.runsParse(null, null, null, null, null, paragraph.getText(), runs, null);					
						
					}
				}
			}
			
			
			List<Table> tables = getTables(report);			
			if (tables != null && tables.size() > 0)
			{
				for (int i=0; i<tables.size(); i++)
				{
					Table table = tables.get(i);
					//解析数据
					msword.report.table.Operate.scavengeSuperfluous(table);
					
				}
			
			}
			
		}
		
	}
	
	/**
	 * 获取报表msword表格列表
	 * @param report
	 * @return
	 */
	public static List<Table> getTables(Report report) {
		List<Table> result = null;
		
		if (report != null 
				&& report.document != null 
				&& report.document.getTables() != null 
				&& report.document.getTables().size() > 0)
		{
			List<Table> tables = new ArrayList<Table>();
			for (int i=0; i<report.document.getTables().size(); i++)
			{
				XWPFTable xwpfTable = report.document.getTables().get(i);
				Table table = new Table();
				table.report = report;
				table.xwpfTable = xwpfTable;
				table.xwpfTableIndex = report.document.getPosOfTable(xwpfTable);
				
				msword.report.table.Operate.analyticStyle(table);			
				
				if (result == null)
				{
					result = new ArrayList<Table>();
				}
				result.add(table);
				

				


			}
			report.tables = result;
			
		}
		
		return result;
	}
	
	/**
	 * 直接获取报表msword表格列表（不做处理）
	 * @param report
	 * @return
	 */
	public static List<Table> getDirectTables(Report report) {
		List<Table> result = null;
		
		if (report != null 
				&& report.document != null 
				&& report.document.getTables() != null 
				&& report.document.getTables().size() > 0)
		{
			List<Table> tables = new ArrayList<Table>();
			for (int i=0; i<report.document.getTables().size(); i++)
			{
				XWPFTable xwpfTable = report.document.getTables().get(i);
				Table table = new Table();
				table.report = report;
				table.xwpfTable = xwpfTable;				
				
				if (result == null)
				{
					result = new ArrayList<Table>();
				}
				result.add(table);


			}
			report.tables = result;
			
		}
		
		return result;
	}
	
	/**
	 * 重载报表msword表格列表
	 * @param report
	 * @return
	 */
	public static List<Table> regetTables(Report report) {
		List<Table> result = null;
		
		if (report != null 
				&& report.tables != null
				&& report.tables.size() > 0
				&& report.document != null 
				&& report.document.getTables() != null 
				&& report.document.getTables().size() > 0)
		{
			for (int i=0; i<report.tables.size(); i++)
			{
				Table table = report.tables.get(i);
				if (table != null)
				{
					XWPFTable xwpfTable = report.document.getTables().get(i);					
					table.xwpfTable = xwpfTable;
					
					if (table.sectionList != null && table.sectionList.size() > 0)
					{
						for (int i1=0; i1<table.sectionList.size(); i1++)
						{
							Section section = table.sectionList.get(i1);
							if (section != null)
							{
								if (section.alt != null && !section.alt.trim().equals(""))
								{
//									section.rptRowTmp = null;
//									section.appendRptRows = null;
//									section.beginIndex = -1;
//									section.endIndex = -1;
								}
								else
								{
									if (section.rptRowTmp != null && section.rptRowTmp.size() > 0)
									{
										for (int i2=0; i2<section.rptRowTmp.size(); i2++)
										{
											TemplateRow templateRow = section.rptRowTmp.get(i2);	
											
											templateRow.tmpRow = table.xwpfTable.getRow(templateRow.index);
											
											
										}								
									}
								}
								
							}
							
									
						}						
					}
					
				}
			}
			
			result = report.tables;
		}
		
		return result;
	}
	
	/**
	 * 载入模板文件
	 * @param report
	 * @return
	 */
	public static XWPFDocument load(Report report) {
		XWPFDocument result = null;
		if (report != null 
				&& report.tmpFilename != null 
				&& !report.tmpFilename.trim().equals(""))
		{
			File tmpFile = new File(report.tmpFilename.trim());
			try {
				if (tmpFile.exists())
				{
					report.is = new FileInputStream(report.tmpFilename.trim()); // 载入文档
					
					ZipSecureFile.setMinInflateRatio(-1.0d);
					report.document = new XWPFDocument(OPCPackage.open(report.is));
					
					result = report.document;
				}
				
			} catch (Exception e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			finally {				
				tmpFile = null;
				
			}
			
		}
		
		return result;
	}
	
	/**
	 * 重载工作文件
	 * @param report
	 * @return
	 */
	public static XWPFDocument reload(Report report) {
		XWPFDocument result = null;
		
		if (report != null)
		{
			if (report.tables != null && report.tables.size() > 0)
			{
				for (int i=0; i<report.tables.size(); i++)
				{
					Table table = report.tables.get(i);
					if (table != null && table.xwpfTable != null)
					{
						table.xwpfTableIndex =  report.document.getPosOfTable(table.xwpfTable);						
					}					
				}
			}		
		}
		
		boolean isSaved = save(report);
		if (isSaved)
		{			
			try {				
				
				report.is = new FileInputStream(report.workFilename.trim()); // 载入工作文档
				ZipSecureFile.setMinInflateRatio(-1.0d);
				report.document = new XWPFDocument(OPCPackage.open(report.is));
				
				regetTables(report);
				
				result = report.document;
				
				
				
				
			} catch (Exception e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			
		}
		
		return result;
	}
	
	/**
	 * 保存工作文件
	 * @param report
	 * @return
	 */
	public static boolean save(Report report) {
		boolean result = false;
		if (report != null 
				&& report.workFilename != null 
				&& !report.workFilename.trim().equals(""))
		{
			if (report.document != null)
			{				
				try {
					
					File workFile = new File(report.workFilename.trim());
					String workPathname = workFile.getParent();
					File workPath = new File(workPathname);
					if ((workPath != null) && (!workPath.isFile()) && !workPath.exists()) 
					{
						workPath.mkdirs();
					}					
					report.os = new FileOutputStream(report.workFilename.trim());
					report.document.write(report.os);
					report.os.close();
					report.is.close();
					
					result = true;
					
				} catch (Exception e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
				
			}
			
		}
		
		return result;
	}
	
	/**
	 * 另保存工作文件
	 * @param report
	 * @param filename 文件带路径名
	 * @return
	 */
	public static boolean saveas(Report report, String filename) {
		boolean result = false;
		if (report != null 
				&& filename != null 
				&& !filename.trim().equals(""))
		{
			if (report.document != null)
			{				
				try {
					
					report.os = new FileOutputStream(filename.trim());
					report.document.write(report.os);
					report.os.close();
					
					result = true;
					
				} catch (Exception e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
				
			}
			
		}
		
		return result;
	}
	
	
	

	
}
