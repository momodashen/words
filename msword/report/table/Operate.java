package msword.report.table;

import java.util.List;

import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import msword.report.Table;

public class Operate {
	
	
	
	/**
	 * 解析表格样式
	 * @param table msword模板中的表格
	 * @return
	 */
	public static Table analyticStyle(Table table) {
				
		if (table != null)
		{
			List<Tag> tagList = msword.report.table.tag.Operate.execute(table);
			table.tagList = tagList;

			List<Section> sectionList = msword.report.table.section.Operate.sectionList(tagList);
			table.sectionList = sectionList;

			List<Section> sectionTree = msword.report.table.section.Operate.tree(sectionList, table.levelCount);
			table.sectionTree = sectionTree;
			
			msword.report.table.section.Operate.execute(sectionTree);
			
			
		}
		
		
		return table;
		
		
	}
	
	/**
	 * 解析表格数据
	 * @param table
	 * @return
	 */
	public static Table analyticData(Table table) {
		
		if (table != null && table.sectionTree != null && table.sectionTree.size() > 0)
		{
			for (int i1=0; i1<table.sectionTree.size(); i1++)
			{
				Section section = table.sectionTree.get(i1);
				msword.report.table.section.Operate.analyticData(section);
				
			}			
		}
		analyticDataOfCorner(table);
		
		return table;
	}
	
	/**
	 * 再替换没有匹配的${}为空字符
	 * @param table
	 * @return
	 */
	public static Table scavengeSuperfluous(Table table) {
		
		if (table != null 
				&& table.xwpfTable != null 
				&& table.xwpfTable.getRows() != null 
				&& table.xwpfTable.getRows().size() > 0)
		{
			for (int i1=0; i1<table.xwpfTable.getRows().size(); i1++)
			{
				XWPFTableRow row = table.xwpfTable.getRow(i1);
				if (row != null && row.getTableCells() != null && row.getTableCells().size() > 0)
				{
					for (int i4=0; i4<row.getTableCells().size(); i4++)
					{
						List<XWPFParagraph> paragraphList = row.getCell(i4).getParagraphs();
						if (paragraphList != null && paragraphList.size() > 0)
						{
							for (int i5=0;i5<paragraphList.size(); i5++)
							{
								XWPFParagraph paragraph = paragraphList.get(i5);
								if (paragraph != null)
								{
									List<XWPFRun> runs = paragraph.getRuns();	
									
									String result = msword.Operate.runsParse(null, null, null, null, null, paragraph.getText(), runs, null);
									
								}
							}
						}
						
					}					
				}
				
			}			
		}
		
		return table;
	}
	
	/**
	 * 解析表格没有被数据区域sectionList处理到的各个角落的数据（不应该在sectionList处理数据前执行该操作）
	 * @param table
	 * @return
	 */
	public static Table analyticDataOfCorner(Table table) {
		
		if (table != null 
				&& table.xwpfTable != null 
				&& table.xwpfTable.getRows() != null 
				&& table.xwpfTable.getRows().size() > 0)
		{
			for (int i=0; i<table.xwpfTable.getRows().size(); i++)
			{
				if (table.cacheIndexList != null && table.cacheIndexList.indexOf(i) > -1)
				{//	已被sectionList处理				
				}
				else
				{
					XWPFTableRow row = table.xwpfTable.getRow(i);					
					if (row != null && row.getTableCells() != null && row.getTableCells().size() > 0)
					{
						for (int i4=0; i4<row.getTableCells().size(); i4++)
						{
							List<XWPFParagraph> paragraphList = row.getCell(i4).getParagraphs();
							if (paragraphList != null && paragraphList.size() > 0)
							{
								for (int i5=0;i5<paragraphList.size(); i5++)
								{
									XWPFParagraph paragraph = paragraphList.get(i5);
									if (paragraph != null)
									{
										List<XWPFRun> runs = paragraph.getRuns();	

										String result = msword.Operate.runsParse(null, null, null, null, null, paragraph.getText(), runs, table.report.varValues);
										
									}
								}
							}
							
							
						}
						
					}
				}
			}
		}
		
		return table;		
	}
	
	/**
	 * 删除开始结束标签
	 * @param table
	 * @return
	 */
	public static Table removeTag(Table table) {
		
		if (table != null && table.sectionTree != null && table.sectionTree.size() > 0)
		{
			for (int i1=0; i1<table.sectionTree.size(); i1++)
			{
				Section section = table.sectionTree.get(i1);				
				msword.report.table.section.Operate.removeTag(section);
				
			}		
		}
		
		return table;
	}
	
	/**
	 * 根据配置合并数据行
	 * @param table
	 * @return
	 */
	public static Table merge(Table table) {
		
		if (table != null && table.sectionTree != null && table.sectionTree.size() > 0)
		{
			for (int i1=0; i1<table.sectionTree.size(); i1++)
			{
				Section section = table.sectionTree.get(i1);				
				msword.report.table.section.Operate.merge(section);
				
			}		
		}
		
		return table;
	}

	

}
