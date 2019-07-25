package msword.report.table.tag;

import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import msword.report.Table;
import msword.report.table.Tag;


public class Operate {
	
	
	/**
	 * 取得规范化的报表表格数据行开始及结束标签列表
	 * @param table msword模板中的表格
	 * @return
	 */
	public static List<Tag> execute(Table table) {
		//抓取标签
		List<Tag> result = grab(table);
		//标识孤立tag
		result = identifierOutlier(result);
		//解析层深及对偶项
		result = leverWeave(result);		
		
		return result;
		
	}	
	
	/**
	 * 抓取报表表格数据行开始及结束标签
	 * @param table msword模板中的表格
	 * @return
	 */
	public static List<Tag> grab(Table table) {
		List<Tag> result = null;
		
		if (table != null && table.xwpfTable != null)
		{
			for (int i=0; i<table.xwpfTable.getRows().size(); i++)
			{
				XWPFTableRow row = table.xwpfTable.getRows().get(i);
				if (row != null && row.getTableCells() != null && row.getTableCells().size() > 0)
				{
					for (int col=0; col<row.getTableCells().size(); col++)
					{
						String text = row.getCell(col).getText();
						
						//解析beginTagIndex和beginTagText
						String regEx = "\\<[\\s]*forEach [^\\>]*\\>";
						Pattern pattern = Pattern.compile(regEx, Pattern.CASE_INSENSITIVE);
						Matcher m = pattern.matcher(text);
						if (m.find()) {
							
							Tag tag = new Tag();
							tag.table = table;						
							tag.type = Type.BeginOf;
							tag.index = i;
							tag.text = m.group(0);
							
							if (result == null)
							{
								result = new ArrayList<Tag>();
							}
							
							result.add(tag);

						} else {
							
						}				
						
						//解析endTagIndex和endTagText
						regEx = "\\<[\\s]*/[\\s]*forEach[\\s]*\\>";
						pattern = Pattern.compile(regEx, Pattern.CASE_INSENSITIVE);
						m = pattern.matcher(text);
						if (m.find()) {
							
							Tag tag = new Tag();
							tag.table = table;	
							tag.type = Type.EndOf;
							tag.index = i;
							tag.text = m.group(0);

							if (result == null)
							{
								result = new ArrayList<Tag>();
							}
							
							result.add(tag);

						} else {
							
						}
						
					}
					
				}
				
			}
		}
		
		return result;
	}
	
	/**
	 * 开始及结束标签列表-标识孤立tag
	 * @param tagList 标签列表
	 * @return
	 */
	public static List<Tag> identifierOutlier(List<Tag> tagList) {
		
		if (tagList != null && tagList.size() > 0)
		{
			List<Tag> tagList_tmp = new ArrayList<Tag>();
			tagList_tmp.addAll(tagList);		
			List<Integer> anchorList = new ArrayList<Integer>();
			tagList_tmp = pairWeave(tagList_tmp, anchorList);
			if (tagList != null && tagList.size() > 0)
			{
				for (int i=tagList.size()-1; i>=0; i--)
				{
					Tag tag = tagList.get(i);
					if (anchorList.contains(tag.index))
					{					
					}
					else
					{//标识孤立tag
						tag.isOutlier = true;
					}
				}
			}
			
		}
		
		return tagList;
	}
	
	/**
	 * 报表表格数据行开始及结束标签配对
	 * @param tagList 标签列表
	 * @param anchorList 锚点列表
	 * @return
	 */
	public static List<Tag> pairWeave(List<Tag> tagList, List<Integer> anchorList) {
		
		if (tagList != null && tagList.size() > 0)
		{
			int index0 = indexOfTagList(tagList, Type.BeginOf);
			int index1 = indexOfTagList(tagList, Type.EndOf);
			
			if (index0 > -1 && index1 > -1)
			{
				if (index1 > index0)
				{
					List<Tag> tagList_tmp = tagList.subList(index0, index1 + 1); //开始标签结束标签之间的子列表
					int index_tmp = lastIndexOfTagList(tagList_tmp, Type.BeginOf);
					index0 = index0 + index_tmp;	
					
					Tag begin = tagList.get(index0);
					Tag end = tagList.get(index1);
					
					if (anchorList != null)
					{
						anchorList.add(begin.index);
						anchorList.add(end.index);
					}
					
					if (tagList.remove(begin))
					{
						if (tagList.remove(end))
						{
							tagList = pairWeave(tagList, anchorList);
						}						
					}
				
				}
				else
				{
					tagList = tagList.subList(index1 + 1, tagList.size()); //结束标签后面的子列表
										
					tagList = pairWeave(tagList, anchorList);
				
				}
			}
			
			
		}		
		
		return tagList;
	}
	
	/**
	 * 解析开始及标签tag的层深及对偶项
	 * @param tagList
	 * @return
	 */
	public static List<Tag> leverWeave(List<Tag> tagList) {
		
		if (tagList != null && tagList.size()  > 0)
		{
			Type lastType = Type.NoneOf;
			int level = 0;
			int levelCount = 0;
			
			for (int i=0; i<tagList.size(); i++)
			{
				Tag tag = tagList.get(i);
				if (tag.isOutlier)
				{//孤立点不处理				
				}
				else
				{
					if (lastType == Type.NoneOf && tag.type == Type.BeginOf)
					{
						level ++;
						levelCount ++;
						tag.lever = level;				
					}			
					if (lastType == Type.BeginOf && tag.type == Type.BeginOf)
					{
						level ++;
						levelCount ++;
						tag.lever = level;				
					}
					if (lastType == Type.EndOf && tag.type == Type.BeginOf)
					{
						tag.lever = level;				
					}
					if (lastType == Type.EndOf && tag.type == Type.EndOf)
					{
						level --;
						tag.lever = level;				
					}
					if (lastType == Type.BeginOf && tag.type == Type.EndOf)
					{
						tag.lever = level;				
					}
					lastType = tag.type;
					
				}
			}
			
			for (int i=1; i<=levelCount; i++) 
			{
				int lastIndex = -1;
				lastType = Type.NoneOf;
				for (int i1=0; i1<tagList.size(); i1++)
				{
					Tag tag = tagList.get(i1);
					if (tag.table != null)
					{
						tag.table.levelCount = levelCount;
					}				
					if (tag.isOutlier)
					{//孤立点不处理				
					}
					else
					{
						if (tag.lever == i)
						{
							if (lastType == Type.BeginOf && tag.type == Type.EndOf)
							{
								tagList.get(lastIndex).pairingIndex = tag.index;
								tag.pairingIndex = tagList.get(lastIndex).index;		
							}
							lastIndex = i1;
							lastType = tag.type;
						}	
					}
				}
			}
		}
		
		return tagList;		
		
	}
		
	/**
	 * 索引第一个开始或结束标签
	 * @param tagList 标签列表
	 * @param type 要索引的开始或结束标签类别   0 开始标签 1 结束标签	
	 */
	public static int indexOfTagList(List<Tag> tagList, Type type) {
		int result = -1;
		
		if (tagList != null && tagList.size() > 0)
		{
			for (int i=0; i<tagList.size(); i++)
			{
				Tag tag = tagList.get(i);
				if (tag != null && tag.type == type)
				{
					result = i;
					break;
				}
			}
		}
		
		return result;
	}
	
	
	/**
	 * 索引最后一个开始或结束标签
	 * @param tagList 标签列表
	 * @param type 要索引的开始或结束标签类别   0 开始标签 1 结束标签	
	 */
	public static int lastIndexOfTagList(List<Tag> tagList, Type type) {
		int result = -1;
		
		if (tagList != null && tagList.size() > 0)
		{
			for (int i=0; i<tagList.size(); i++)
			{
				Tag tag = tagList.get(i);
				if (tag != null && tag.type == type)
				{
					result = i;
				}
			}
		}
		
		return result;
	}

}
