package msword.report.root;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import msword.Report;
import msword.report.Root;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

public class Operate {
	
	/**
	 * 取得模板的root数据标签列表
	 * @param report 
	 * @return
	 */
	public static List<Root> execute(Report report) {
		List<Root> result = null;
		if (report != null 
				&& report.document != null 
				&& report.document.getTables() != null 
				&& report.document.getTables().size() > 0)
		{
			for (int i=0; i<report.document.getTables().size(); i++)
			{
				XWPFTable xwpfTable = report.document.getTables().get(i);
				List<Root> list = grab(xwpfTable);
				if (list != null && list.size() > 0)
				{
					if (result == null)
					{
						result = new ArrayList<Root>();
					}
					result.addAll(list);
				}				
			}
			
//			result = textParse(result, report.varValues);
			
		}
		
		report.rootList = result;
		
		return result;
		
	}		
	
	/**
	 * root数据的定义中的sql文本作变量替换
	 * @param rootList root数据的定义列表
	 * @param varValues 存放变量占位符及对应值的map
	 * @return
	 */
	public static List<Root> textParse(List<Root> rootList, Map<String, Object> varValues) {
		if (rootList != null 
				&& rootList.size() > 0 
				&& varValues != null 
				&& !varValues.isEmpty())
		{
			for (int i=0; i<rootList.size(); i++)
			{
				Root root = rootList.get(i);
				if (root != null 
						&& root.sql != null 
						&& !root.sql.trim().equals(""))
				{
					String text = root.sql;
					text = msword.Operate.textParse(null, null, null, null, null, text, varValues);
					root.sql = text;
				}
				
			}
			
		}
		
		return rootList;
	}
	
	
	/**
	 * 抓取报表root数据标签
	 * @param xwpfTable msword模板中的表格
	 * @return
	 */
	public static List<Root> grab(XWPFTable xwpfTable) {
		List<Root> result = null;
		
		if (xwpfTable != null)
		{
			for (int i=0; i<xwpfTable.getRows().size(); i++)
			{
				XWPFTableRow row = xwpfTable.getRows().get(i);
				if (row != null && row.getTableCells() != null && row.getTableCells().size() > 0)
				{
					String text = row.getCell(0).getText();
					
					//解析beginTagIndex和beginTagText
					String regEx = "\\<[\\s]*root [^\\>]*\\>";
					Pattern pattern = Pattern.compile(regEx, Pattern.CASE_INSENSITIVE);
					Matcher m = pattern.matcher(text);
					if (m.find()) {
						
						Root root = new Root();
						root.xwpfTable = xwpfTable;						
						root.index = i;
						root.text = m.group(0);						
						
						//解析SQL
						regEx = "SQL[\\s]*=[\\s]*\\[[\\s]*([^\\]])*\\]";
						pattern = Pattern.compile(regEx, Pattern.CASE_INSENSITIVE);
						m = pattern.matcher(root.text);
						if (m.find()) {
							String tmp  = m.group(0);
							
							System.out.println(tmp);
							
							int b = tmp.indexOf("[");
							int e = tmp.lastIndexOf("]");
							root.sql = tmp.substring(b+1, e);
							
							System.out.println(root.sql);

						} else {
							
						}				
						
						//解析placeholder
						regEx = "var[\\s]*=[\\s]*[\"“”][\\s]*[^\"“”]*[\"“”]";
						pattern = Pattern.compile(regEx, Pattern.CASE_INSENSITIVE);
						m = pattern.matcher(root.text);
						if (m.find()) {
							//String[] tmpArr = m.group(0).split("=");
							//root.placeholder = tmpArr[1].trim().replaceAll("[\"“”\\s　 ]*", "");
							String tmpStr = m.group(0);
							tmpStr = tmpStr.replaceFirst("var[\\s]*=[\\s]*[\"“”]", "");							
							root.placeholder = tmpStr.replaceFirst("[\"“”]$", "");		

						} else {
							regEx = "var[\\s]*=[\\s]*[^\\s\"“”]*[\\s]*";
							pattern = Pattern.compile(regEx, Pattern.CASE_INSENSITIVE);
							m = pattern.matcher(root.text);
							if (m.find()) {
								//String[] tmpArr = m.group(0).split("=");
								//root.placeholder = tmpArr[1].trim().replaceAll("[\"“”\\s　 ]*", "");
								String tmpStr = m.group(0);
								tmpStr = tmpStr.replaceFirst("var[\\s]*=[\\s]*", "");									
								root.placeholder = tmpStr;

							} else {								
							}

						}							
						
						//解析dateFormat
						regEx = "dateFormat[\\s]*=[\\s]*[\"“”][\\s]*[^\"“”]*[\"“”]";
						pattern = Pattern.compile(regEx, Pattern.CASE_INSENSITIVE);
						m = pattern.matcher(root.text);
						if (m.find()) {
							//String[] tmpArr = m.group(0).split("=");
							//root.dateFormat = tmpArr[1].trim().replaceAll("[\"“”\\s　 ]*", "");
							String tmpStr = m.group(0);
							tmpStr = tmpStr.replaceFirst("dateFormat[\\s]*=[\\s]*[\"“”]", "");							
							root.dateFormat = tmpStr.replaceFirst("[\"“”]$", "");

						} else {
							regEx = "dateFormat[\\s]*=[\\s]*[^\\s\"“”]*[\\s]*";
							pattern = Pattern.compile(regEx, Pattern.CASE_INSENSITIVE);
							m = pattern.matcher(root.text);
							if (m.find()) {
								//String[] tmpArr = m.group(0).split("=");
								//root.dateFormat = tmpArr[1].trim().replaceAll("[\"“”\\s　 ]*", "");
								String tmpStr = m.group(0);
								tmpStr = tmpStr.replaceFirst("dateFormat[\\s]*=[\\s]*", "");									
								root.dateFormat = tmpStr;

							} else {								
							}

						}	
						
						//解析datetimeFormat
						regEx = "datetimeFormat[\\s]*=[\\s]*[\"“”][\\s]*[^\"“”]*[\"“”]";
						pattern = Pattern.compile(regEx, Pattern.CASE_INSENSITIVE);
						m = pattern.matcher(root.text);
						if (m.find()) {
							//String[] tmpArr = m.group(0).split("=");
							//root.datetimeFormat = tmpArr[1].trim().replaceAll("[\"“”\\s　 ]*", "");
							String tmpStr = m.group(0);
							tmpStr = tmpStr.replaceFirst("datetimeFormat[\\s]*=[\\s]*[\"“”]", "");							
							root.datetimeFormat = tmpStr.replaceFirst("[\"“”]$", "");

						} else {
							regEx = "datetimeFormat[\\s]*=[\\s]*[^\\s\"“”]*[\\s]*";
							pattern = Pattern.compile(regEx, Pattern.CASE_INSENSITIVE);
							m = pattern.matcher(root.text);
							if (m.find()) {
								//String[] tmpArr = m.group(0).split("=");
								//root.datetimeFormat = tmpArr[1].trim().replaceAll("[\"“”\\s　 ]*", "");
								String tmpStr = m.group(0);
								tmpStr = tmpStr.replaceFirst("datetimeFormat[\\s]*=[\\s]*", "");									
								root.datetimeFormat = tmpStr;

							} else {								
							}

						}	
						
						//解析fieldsOfDatetime
						regEx = "fieldsOfDatetime[\\s]*=[\\s]*[\"“”][\\s]*[^\"“”]*[\"“”]";
						pattern = Pattern.compile(regEx, Pattern.CASE_INSENSITIVE);
						m = pattern.matcher(root.text);
						if (m.find()) {
							//String[] tmpArr = m.group(0).split("=");
							//root.fieldsOfDatetime = tmpArr[1].trim().replaceAll("[\"“”\\s　 ]*", "");
							String tmpStr = m.group(0);
							tmpStr = tmpStr.replaceFirst("fieldsOfDatetime[\\s]*=[\\s]*[\"“”]", "");							
							root.fieldsOfDatetime = tmpStr.replaceFirst("[\"“”]$", "");

						} else {
							regEx = "fieldsOfDatetime[\\s]*=[\\s]*[^\\s\"“”]*[\\s]*";
							pattern = Pattern.compile(regEx, Pattern.CASE_INSENSITIVE);
							m = pattern.matcher(root.text);
							if (m.find()) {
								//String[] tmpArr = m.group(0).split("=");
								//root.fieldsOfDatetime = tmpArr[1].trim().replaceAll("[\"“”\\s　 ]*", "");
								String tmpStr = m.group(0);
								tmpStr = tmpStr.replaceFirst("fieldsOfDatetime[\\s]*=[\\s]*", "");									
								root.fieldsOfDatetime = tmpStr;

							} else {								
							}

						}	
						
						
						//解析isomer
						String isomerStr = "";
						regEx = "isomer[\\s]*=[\\s]*[\"“”][\\s]*\\[[\\s]*[\\s]*\\{[\\s]*[^\"“”]*\\}[\\s]*\\][\\s]*[\"“”]";
						pattern = Pattern.compile(regEx, Pattern.CASE_INSENSITIVE);
						m = pattern.matcher(root.text);
						if (m.find()) {
							String tmpStr = m.group(0);
							tmpStr = tmpStr.replaceFirst("isomer[\\s]*=[\\s]*[\"“”]", "");							
							isomerStr = tmpStr.replaceFirst("[\"“”]$", "");

						} else {
							regEx = "isomer[\\s]*=[\\s]*\\[[\\s]*[\\s]*\\{[\\s]*[^\"“”]*\\}[\\s]*\\][\\s]*";
							pattern = Pattern.compile(regEx, Pattern.CASE_INSENSITIVE);
							m = pattern.matcher(root.text);
							if (m.find()) {
								String tmpStr = m.group(0);
								tmpStr = tmpStr.replaceFirst("isomer[\\s]*=[\\s]*", "");									
								isomerStr = tmpStr;

							} else {								
							}

						}	
						try {
							root.isomer = com.alibaba.fastjson.JSONArray.parseArray(isomerStr); //异构样式配置								
						} catch (Exception e) {
							root.isomer = null;
						}
						
						
						
						
						if (result == null)
						{
							result = new ArrayList<Root>();
						}
						
						result.add(root);
						int selfIndex  = result.size() - 1;
						root.selfIndex = selfIndex;

					} else {
						
					}
					
				}
				
			}
		}
		
		return result;
	}
	
	/**
	 * 删除root数据标签
	 * @param rootList
	 * @return
	 */
	public static Report removeTag(Report report) {
		
		if (report != null 
				&& report.document != null  
				&& report.rootList != null 
				&& report.rootList.size() > 0)
		{
			for (int j=0; j<report.rootList.size(); j++)
			{
				Root root = report.rootList.get(j);
				//删除数据区域中开始标签
				removeRow(root, report.rootList);				
				if (root.xwpfTable != null)
				{
					if (root.xwpfTable.getRows() != null && root.xwpfTable.getRows().size() > 0) { } else
					{
						int tIndex = report.document.getPosOfTable(root.xwpfTable);
						report.document.removeBodyElement(tIndex);
					}
				}
			}
		}
		
		return report;
	}

	/**
	 * 删除root数据标签某一行
	 * @param root 要删除行所在的root数据标签
	 * @param rootList root数据标签列表
	 */
	public static void removeRow(Root root, List<Root> rootList) {
		if (root != null 
				&& root.xwpfTable != null)
		{
			if (root.index > -1)
			{
				root.xwpfTable.removeRow(root.index);
				root.index = -1;
				
				for (int j=root.selfIndex+1; j<rootList.size(); j++)
				{
					Root tar = rootList.get(j);
					if (tar.xwpfTable == root.xwpfTable)
					{
						tar.index --;						
					}
					
				}
				
				
			}
			
		}
	}
		
}
