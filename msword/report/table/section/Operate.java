package msword.report.table.section;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.hibernate.transform.Transformers;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTVMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;

import msword.report.Table;
import msword.report.table.Section;
import msword.report.table.Tag;
import msword.report.table.tag.Type;

public class Operate {

	
	/**
	 * 
	 * @param tagList 标签列表
	 * @return
	*/
	public static List<Section> execute(List<Section> sectionList) {
		
		if (sectionList != null && sectionList.size() > 0)
		{
			for (int i=0; i<sectionList.size(); i++)
			{
				Section section = sectionList.get(i);
				
				monolayerStyle(section);
			}
			
		}	
		
		return sectionList;
		
	}
	
	/**
	 * 只处理单层数据行样式
	 * @param section
	 * @return
	 */
	public static Section monolayerStyle(Section section) {
		
		if (section != null 
				&& section.table != null 
				&& section.table.report != null 				
				&& section.table.report.varValues != null
				&& section.rowPlaceholder != null 
				&& !section.rowPlaceholder.trim().equals(""))
		{
			List<SectionRow> appendRptRows = new ArrayList<SectionRow>(); //根据数据源的记录数添加的表格行索引列表

			Map<String, Object> varValues = section.table.report.varValues;
			if (varValues.isEmpty())
			{									
			}
			else
			{
				if (section.rptRowTmp != null && section.rptRowTmp.size() > 0)
				{		
					List items = null;
					if (section.dataSourceKey != null 
							&& !section.dataSourceKey.trim().equals(""))
					{
						Object dataSource = varValues.get(section.dataSourceKey.trim());
						if (dataSource != null)
						{
							if (dataSource instanceof List)
							{
								items = (List)dataSource;
							}
						}
					}
					if (items == null)
					{
						if (section.sql != null 
								&& !section.sql.trim().equals(""))
						{
							items = section.table.report.smartBaseDao.getSqlQuery(section.sql).setResultTransformer(Transformers.ALIAS_TO_ENTITY_MAP).list();
						}
					}
					if (items != null && items.size() > 0)
					{
						//按照报表行模板添加报表行
						
						for (int i1=0; i1<items.size(); i1++)
						{
							Object rowSource = items.get(i1);
							if (rowSource != null)
							{
								if (rowSource instanceof Map)
								{
									Map item = (Map)rowSource;
									
									List<Integer> rptRow = new ArrayList<Integer>();
									
									for (int i2=0; i2<section.rptRowTmp.size(); i2++)
									{
										TemplateRow templateRow = section.rptRowTmp.get(i2);	
										if (templateRow != null && templateRow.tmpRow != null)
										{
											if (i1 == 0)
											{
												rptRow.add(templateRow.index);														
											}
											else
											{
												boolean issucc = section.table.xwpfTable.addRow(templateRow.tmpRow, section.endIndex);
												if (issucc)
												{
													rptRow.add(section.endIndex);
													
													section.endIndex ++;
												}
											}
											
										}
										
									}
									
									SectionRow sectionRow = new SectionRow();
									sectionRow.rptRow = rptRow;
									sectionRow.varValues = item;		
									
									appendRptRows.add(sectionRow);
									
									
								}	
							}															
							
						}								
					
					}	
					else
					{
						if (section.alt != null && !section.alt.trim().equals(""))
						{
							if (section.table != null && section.table.xwpfTable != null)
							{						
								/*
								int size = section.table.xwpfTable.getRows().size();
								for (int pos=0; pos<size-1; pos++)
								{
									section.table.xwpfTable.removeRow(0);										
								}			
								XWPFTableRow row = section.table.xwpfTable.getRow(0);	
								*/	
								if (section.colTitleIndex > -1)
								{
									removeRow(section, section.colTitleIndex, section.table.sectionList);
								}
								if (section.beginIndex > -1)
								{
									removeRow(section, section.beginIndex, section.table.sectionList);
								}
								int size = section.rptRowTmp.size();
								for (int i2=0; i2<size; i2++)
								{
									TemplateRow templateRow = section.rptRowTmp.get(i2);	
									if (templateRow != null && templateRow.index > -1)
									{
										removeRow(section, templateRow.index, section.table.sectionList);
									}
								}								
								XWPFTableRow row = section.table.xwpfTable.getRow(section.endIndex);	
								
								
								if (row.getTableICells() != null && row.getTableICells().size() > 0)
								{
									mergeCellHorizontally(section.table.xwpfTable, 0, 0, row.getTableICells().size()-1);
									if (row.getCell(0).getParagraphs() != null 
											&& row.getCell(0).getParagraphs().size() > 0 
											&& row.getCell(0).getParagraphs().get(0) != null 
											&& row.getCell(0).getParagraphs().get(0).getRuns() != null 
											&& row.getCell(0).getParagraphs().get(0).getRuns().size() > 0
											&& row.getCell(0).getParagraphs().get(0).getRuns().get(0) != null)
									{
										row.getCell(0).getParagraphs().get(0).getRuns().get(0).setText(section.alt,0);		
										for (int i4=1; i4<row.getCell(0).getParagraphs().get(0).getRuns().size(); i4++)
										{
											row.getCell(0).getParagraphs().get(0).getRuns().get(i4).setText("",0);		
										}
									}
								}
								else
								{
									row.createCell().setText(section.alt);		
								}
								
							}
							
							
						}
						
					}
				
				}
			}	
			
			if (section != null)
			{
				section.appendRptRows = appendRptRows;
				
				if (appendRptRows.size() > 0 
						&& section.table.sectionList != null 
						&& section.table.sectionList.size() > 0)
				{	
					
					for (int j=section.selfIndex+1; j<section.table.sectionList.size(); j++)
					{
						Section tar = section.table.sectionList.get(j);
						if (tar != null && tar.rptRowTmp != null && tar.rptRowTmp.size() > 0)
						{
							tar.colTitleIndex += (appendRptRows.size() - 1) * section.rptRowTmp.size(); 
							tar.beginIndex += (appendRptRows.size() - 1) * section.rptRowTmp.size();
							tar.endIndex += (appendRptRows.size() - 1) * section.rptRowTmp.size();				
							
							for (int i2=0; i2<tar.rptRowTmp.size(); i2++)
							{
								TemplateRow templateRow = tar.rptRowTmp.get(i2);	
								if (templateRow != null)
								{
									templateRow.index += (appendRptRows.size() - 1) * section.rptRowTmp.size();
								}
								
							}
						

						}
						
					}
				
				}
			}			
		
		}
		
		return section;
	}	 
	
	/**
	 * 解析数据区域数据
	 * @param section
	 * @return
	 */
	public static Section analyticData(Section section) {
		
		if (section != null 
				&& section.table != null 
				&& section.table.report != null
				&& section.table.xwpfTable != null 
				&& section.appendRptRows != null 
				&& section.appendRptRows.size() > 0)
		{			
			for (int i2=0; i2<section.appendRptRows.size(); i2++)
			{
				SectionRow sectonRow = section.appendRptRows.get(i2);
				if (sectonRow != null 
						&& sectonRow.varValues != null 
						&& sectonRow.rptRow != null 
						&& sectonRow.rptRow.size() > 0)
				{
					for (int i3=0; i3<sectonRow.rptRow.size(); i3++)
					{
						int rowIndex = sectonRow.rptRow.get(i3);				
						XWPFTableRow row = section.table.xwpfTable.getRows().get(rowIndex);
						
						if (section.table.cacheIndexList == null)
						{
							section.table.cacheIndexList = new ArrayList<Integer>();
						}
						section.table.cacheIndexList.add(rowIndex);
						
						if (row != null && row.getTableCells() != null && row.getTableCells().size() > 0)
						{
							for (int i4=0; i4<row.getTableCells().size(); i4++)
							{
								List<XWPFParagraph> paragraphList = row.getCell(i4).getParagraphs();
								if (paragraphList != null && paragraphList.size() > 0)
								{
									if (section.colOfSerialNumber > -1 && section.colOfSerialNumber == i4 && i3 == 0)
									{//有输出序号配置	
										paragraphList.get(0).createRun().setText(String.valueOf(i2 + 1), 0);
									}
									else
									{
										for (int i5=0;i5<paragraphList.size(); i5++)
										{
											XWPFParagraph paragraph = paragraphList.get(i5);
											if (paragraph != null)
											{
												List<XWPFRun> runs = paragraph.getRuns();	
												
												msword.Operate.pictureAnaly(sectonRow.varValues, section.rowPlaceholder, runs);

												String result = msword.Operate.runsParse(section.rowPlaceholder, section.dateFormat, section.datetimeFormat, section.fieldsOfDatetime, section.isomer, paragraph.getText(), runs, sectonRow.varValues);
												
												
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
		
		
		return section;
	}
	
	/**
	 * 删除开始结束标签
	 * @param section
	 * @return
	 */
	public static Section removeTag(Section section) {
		
		if (section != null)
		{
			boolean doDelete = true;
			
			if (section.alt != null && !section.alt.trim().equals(""))
			{
				if (section.appendRptRows != null && section.appendRptRows.size() > 0)
				{					
				}
				else
				{
					doDelete = false;
				}				
			}
			if (doDelete)
			{
				//删除数据区域中开始标签
				removeRow(section, section.beginIndex, section.table.sectionList);
				//删除数据区域中结束标签
				removeRow(section, section.endIndex, section.table.sectionList);				
			}
		}
		
		return section;
	}
	
	/**
	 * 根据配置合并指定列的数据行（合并行必须在开始结束标签删除前执行）
	 * @param section
	 * @return
	 */
	public static Section merge(Section section) {
		
		if (section != null 
				&& section.table != null
				&& section.table.xwpfTable != null
				&& section.table.xwpfTable.getRows() != null
				&& section.table.xwpfTable.getRows().size() > 0
				&& section.beginIndex > -1
				&& section.endIndex > - 1
				&& section.fieldsOfMergerow != null 
				&& !section.fieldsOfMergerow.trim().equals(""))
		{
			boolean doMerge = true;
			
			if (section.alt != null && !section.alt.trim().equals(""))
			{
				if (section.appendRptRows != null && section.appendRptRows.size() > 0)
				{					
				}
				else
				{
					doMerge = false;
				}				
			}
			
			if (doMerge)
			{
				String[] fieldsOf = section.fieldsOfMergerow.trim().split(",");
				if (fieldsOf != null && fieldsOf.length > 0) { } else
				{
					fieldsOf = section.fieldsOfMergerow.trim().split("，");
				}
				if (fieldsOf != null && fieldsOf.length > 0) { } else
				{
					fieldsOf = section.fieldsOfMergerow.trim().split(";");
				}
				if (fieldsOf != null && fieldsOf.length > 0) { } else
				{
					fieldsOf = section.fieldsOfMergerow.trim().split("；");
				}
				if (section.rptRowTmp != null && section.rptRowTmp.size() == 1)
				{
					//只有数据区域是单行模式才合并行
					TemplateRow templateRow = section.rptRowTmp.get(0);	
					if (templateRow != null 
							&& templateRow.columnList != null 
							&& templateRow.columnList.size() > 0)
					{
						//检索出需要合并数据行的那些列
						List<Column> colOfMergeList = new ArrayList<Column>();			
						for (int i=0; i<fieldsOf.length; i++)
						{
							String field = fieldsOf[i];
							if (field != null && !field.trim().equals(""))
							{
								for (int i4=0; i4<templateRow.columnList.size(); i4++)
								{
									Column column = templateRow.columnList.get(i4);
									if (column != null)
									{
										if (column.xwpfTableList != null && column.xwpfTableList.size() > 0) {/*内含嵌套表格不处理*/ } else
										{
											String text = column.text;								
											if (text != null && !text.trim().equals(""))
											{
												String regEx = "\\$\\{[\\s　 ]*" + field.trim() + "[\\s　 ]*\\}";
												if (section.rowPlaceholder != null && !section.rowPlaceholder.trim().equals(""))
												{
													regEx = "\\$\\{[\\s　 ]*" + section.rowPlaceholder.trim() + "\\.[\\s　 ]*" + field.trim() + "[\\s　 ]*\\}";
												}
												
												Pattern pattern = Pattern.compile(regEx, Pattern.CASE_INSENSITIVE);
												Matcher m = pattern.matcher(text);
												if (m.find()) {
													
													Column colOfMerge = new Column();
													colOfMerge.field = field;
													colOfMerge.index = column.index;
													
													colOfMergeList.add(colOfMerge);
													
												}
												
											}
										}
										
									}
									
									
								}
								
							}						
						}
						
						if (colOfMergeList != null && colOfMergeList.size() > 0)
						{
							for (int i=0; i<colOfMergeList.size(); i++)
							{
								Column colOfMerge = colOfMergeList.get(i);
								if (colOfMerge != null && colOfMerge.index > -1)
								{
									int col = colOfMerge.index;
									
									List<IntBE> rangeList = new ArrayList<IntBE>();								
									int tmpBIndex = -1;
									int tmpEIndex = -1;
									String tmpText = "";
									for (int i1=section.beginIndex+1; i1<section.endIndex; i1++)
									{
										String text = section.table.xwpfTable.getRow(i1).getCell(col).getText();
										if (text != null && !text.trim().equals("")) { } else
										{
											text = "";
										}
										if (tmpText != null && !tmpText.trim().equals("")) { } else
										{
											tmpText = "";
										}

										if (tmpText.trim().equals(text.trim()))
										{//
											tmpEIndex = i1;
										}
										else
										{//
											if (tmpEIndex == -1) { } else
											{//一个不同的内容的开始 记录保存在内存中的begin - end 对
												if (tmpBIndex > -1 
														&& tmpEIndex > -1 
														&& (tmpEIndex - tmpBIndex) > 0)  // (tmpEIndex - tmpBIndex) > 1
												{
													IntBE be = new IntBE();
													be.b = tmpBIndex;
													be.e = tmpEIndex;
													be.text = tmpText;
													
													rangeList.add(be);	
													
												}																							
											}
											
											tmpBIndex = i1;
											tmpEIndex = i1;
										}
									
										tmpText = text;
										
									}
									if (tmpBIndex == tmpEIndex)
									{									
									}
									else if (tmpBIndex > -1 
											&& tmpEIndex > -1 
											&& (tmpEIndex - tmpBIndex) > 0) //(tmpEIndex - tmpBIndex) > 1
									{
										IntBE be = new IntBE();
										be.b = tmpBIndex;
										be.e = tmpEIndex;
										be.text = tmpText;
										
										rangeList.add(be);	
									}
									
									if (rangeList != null && rangeList.size() > 0)
									{
										for (int i2=0; i2<rangeList.size(); i2++)
										{
											IntBE be = rangeList.get(i2);
											if (be != null 
													&& be.b > -1 
													&& be.e > -1 
													&& (be.e - be.b) > 0) //(be.e - be.b) > 1
											{
												mergeCellVertically(section.table.xwpfTable, col, be.b, be.e);
												
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
		
		return section;
	}
	
	/**
	 * 数据区域列表转换位树结构
	 * @param table
	 * @return
	 */
	public static List<Section> tree(List<Section> sectionList, int levelCount) {
		List<Section> result = null;
		if (sectionList != null && sectionList.size() > 0)
		{
			for (int level=1; level<=levelCount; level++)
			{
				for (int i1=0; i1<sectionList.size(); i1++)
				{
					Section section = sectionList.get(i1);
					if (level == 1)
					{
						if (section.lever == 1)
						{
							if (result == null)
							{
								result = new ArrayList<Section>();
							}						
							result.add(section);
						}						
					}
					else
					{
						if (section.lever == level)
						{
							parent(section, i1, sectionList);
						}						
					}					
				}
				
			}
		}		
		
		return result;
	}
	
	/**
	 * 找数据区域父数据区域节点
	 * @param section 数据区域
	 * @param level 层次
	 */
	public static Section parent(Section section, int fromIndex, List<Section> sectionList) {
		Section result = null;
		
		if (section != null 
				&& sectionList != null 
				&& sectionList.size() > 0)
		{
			for (int i=fromIndex-1; i>=0; i--)
			{
				Section tar = sectionList.get(i);
				if (tar.lever == section.lever - 1)
				{
					if (tar.children == null)
					{
						tar.children = new ArrayList<Section>();
					}
					tar.children.add(section);
					
					result = tar;
					
					break;
				}
			}
			
		}
		
		return result;
		
	}
	
	/**
	 * 把“报表表格数据行开始及结束标签”列表解析成“数据区域”列表
	 * @param tagList
	 * @return
	 */
	public static List<Section> sectionList(List<Tag> tagList) {
		List<Section> result = null;
		if (tagList != null && tagList.size() > 0)
		{
			for (int i=0; i<tagList.size(); i++)
			{
				Tag tag = tagList.get(i);
				if (tag != null)
				{
					Section section = analysis(tag);
					if (section != null)
					{
						if (result == null) 
						{
							result = new ArrayList<Section>();
						}
						result.add(section);
						int selfIndex  = result.size() - 1;
						section.selfIndex = selfIndex;
						
					}
					
				}				
			}			
		}	
		
		return result;
	}
	
	/**
	 * 从报表表格数据行开始及结束标签中抓取数据区域
	 * @param tagList 标签列表
	 * @return
	 */
	public static Section analysis(Tag tag) {
		Section result = null;
		
		if (tag != null)
		{
			if (tag.isOutlier)
			{//孤立点不处理						
			}
			else
			{
				if (tag.type == Type.BeginOf)
				{//只处理开始标签就可以了
					Table table = tag.table; //所属的msword表 
					int lever = tag.lever; //层深
					int beginIndex = tag.index; //开始索引
					int endIndex = tag.pairingIndex; ////结束索引
					String attributeText = tag.text; //属性文本
					int colTitleIndex = beginIndex -1; //列标题行索引
					String alt = ""; //当数据源无数据时,用该文本替换
					String sql = ""; //取数据的sql语句
					String dataSourceKey = "items"; //数据源的名称
					String rowPlaceholder = "item"; //数据行占位符
					int colOfSerialNumber = -1; //序号所在的列索引 索引以0开始 第一列为0
					String dateFormat = "yyyy-MM-dd"; //日期格式字符串 需要符合java.text.SampleDateFormat规范
					String datetimeFormat = "yyyy-MM-dd HH:mm:ss"; //日期时间格式字符串 需要符合java.text.SimpleDateFormat规范
					String fieldsOfDatetime = ""; //日期时间字段名称序列,多个已英文逗号分开， 形如    deptname,username 
					String fieldsOfMergerow = ""; //需要合并行数据的列字段名称序列,多个已英文逗号分开， 形如    deptname,username 
					
					String isomer = ""; //异构样式配置 json对象 字符串格式，形如isomer="[{field:'x13',it:[{logicalexpr:'(? + ${item.x14}) eq 2',style:{color:'97FFFF'}}]},{field:'x15',it:[{logicalexpr:'? eq {sysdate}',style:{color:'97FFFF'}}]}]"

					List<TemplateRow> rptRowTmp = new ArrayList<TemplateRow>(); //行模板	
					
					if (beginIndex > -1 && endIndex > -1)
					{
						//解析colTitleIndex
						String regEx = "colTitleIndex[\\s]*=[\\s]*[\"“”][\\s]*[0-9]*[\"“”]";
						Pattern pattern = Pattern.compile(regEx, Pattern.CASE_INSENSITIVE);
						Matcher m = pattern.matcher(attributeText);
						if (m.find()) {
							//String[] tmpArr = m.group(0).split("=");
							//String colTitleIndexStr = tmpArr[1].trim().replaceAll("[\"“”\\s　 ]*", "");
							String tmpStr = m.group(0);
							tmpStr = tmpStr.replaceFirst("colTitleIndex[\\s]*=[\\s]*[\"“”]", "");							
							String colTitleIndexStr = tmpStr.replaceFirst("[\"“”]$", "");
							if (colTitleIndexStr != null && !colTitleIndexStr.trim().equals(""))
							{
								colTitleIndex = Integer.valueOf(colTitleIndexStr.trim());
							}

						} else {
							regEx = "colTitleIndex[\\s]*=[\\s]*[\"“”][\\s]*[0-9]*[\"“”]";
							pattern = Pattern.compile(regEx, Pattern.CASE_INSENSITIVE);
							m = pattern.matcher(attributeText);
							if (m.find()) {
								//String[] tmpArr = m.group(0).split("=");
								//String colTitleIndexStr = tmpArr[1].trim().replaceAll("[\"“”\\s　 ]*", "");
								String tmpStr = m.group(0);
								tmpStr = tmpStr.replaceFirst("colTitleIndex[\\s]*=[\\s]*", "");									
								String colTitleIndexStr = tmpStr;
								if (colTitleIndexStr != null && !colTitleIndexStr.trim().equals(""))
								{
									colTitleIndex = Integer.valueOf(colTitleIndexStr.trim());
								}

							} else {
							}

						}	
						
						//解析alt
						regEx = "alt[\\s]*=[\\s]*[\"“”][\\s]*[^\"“”]*[\"“”]";
						pattern = Pattern.compile(regEx, Pattern.CASE_INSENSITIVE);
						m = pattern.matcher(attributeText);
						if (m.find()) {
							//String[] tmpArr = m.group(0).split("=");
							//alt = tmpArr[1].trim().replaceAll("[\"“”\\s　 ]*", "");
							String tmpStr = m.group(0);
							tmpStr = tmpStr.replaceFirst("alt[\\s]*=[\\s]*[\"“”]", "");							
							alt = tmpStr.replaceFirst("[\"“”]$", "");

						} else {
							regEx = "alt[\\s]*=[\\s]*[^\\s\"“”]*[\\s]*";
							pattern = Pattern.compile(regEx, Pattern.CASE_INSENSITIVE);
							m = pattern.matcher(attributeText);
							if (m.find()) {
								//String[] tmpArr = m.group(0).split("=");
								//alt = tmpArr[1].trim().replaceAll("[\"“”\\s　 ]*", "");
								String tmpStr = m.group(0);
								tmpStr = tmpStr.replaceFirst("alt[\\s]*=[\\s]*", "");									
								alt = tmpStr;

							} else {
							}

						}	

						//解析SQL
						regEx = "SQL[\\s]*=[\\s]*\\[[\\s]*([^\\]])*\\]";
						pattern = Pattern.compile(regEx, Pattern.CASE_INSENSITIVE);
						m = pattern.matcher(attributeText);
						if (m.find()) {
							String tmp  = m.group(0);
							
							int b = tmp.indexOf("[");
							int e = tmp.lastIndexOf("]");
							sql = tmp.substring(b+1, e);

						} else {
							
						}				
						
						//解析dataSourceKey
						regEx = "items[\\s]*=[\\s]*[\"“”][\\s]*[^\"“”]*[\"“”]";
						pattern = Pattern.compile(regEx, Pattern.CASE_INSENSITIVE);
						m = pattern.matcher(attributeText);
						if (m.find()) {
							//String[] tmpArr = m.group(0).split("=");
							//dataSourceKey = tmpArr[1].trim().replaceAll("[\"“”\\s　 ]*", "");
							String tmpStr = m.group(0);
							tmpStr = tmpStr.replaceFirst("items[\\s]*=[\\s]*[\"“”]", "");							
							dataSourceKey = tmpStr.replaceFirst("[\"“”]$", "");

						} else {
							regEx = "items[\\s]*=[\\s]*[^\\s\"“”]*[\\s]*";
							pattern = Pattern.compile(regEx, Pattern.CASE_INSENSITIVE);
							m = pattern.matcher(attributeText);
							if (m.find()) {
								//String[] tmpArr = m.group(0).split("=");
								//dataSourceKey = tmpArr[1].trim().replaceAll("[\"“”\\s　 ]*", "");
								String tmpStr = m.group(0);
								tmpStr = tmpStr.replaceFirst("items[\\s]*=[\\s]*", "");									
								dataSourceKey = tmpStr;

							} else {
							}

						}	
						
						//解析rowPlaceholder
						regEx = "var[\\s]*=[\\s]*[\"“”][\\s]*[^\"“”]*[\"“”]";
						pattern = Pattern.compile(regEx, Pattern.CASE_INSENSITIVE);
						m = pattern.matcher(attributeText);
						if (m.find()) {
							//String[] tmpArr = m.group(0).split("=");
							//rowPlaceholder = tmpArr[1].trim().replaceAll("[\"“”\\s　 ]*", "");
							String tmpStr = m.group(0);
							tmpStr = tmpStr.replaceFirst("var[\\s]*=[\\s]*[\"“”]", "");							
							rowPlaceholder = tmpStr.replaceFirst("[\"“”]$", "");	

						} else {
							regEx = "var[\\s]*=[\\s]*[^\\s\"“”]*[\\s]*";
							pattern = Pattern.compile(regEx, Pattern.CASE_INSENSITIVE);
							m = pattern.matcher(attributeText);
							if (m.find()) {
								//String[] tmpArr = m.group(0).split("=");
								//rowPlaceholder = tmpArr[1].trim().replaceAll("[\"“”\\s　 ]*", "");
								String tmpStr = m.group(0);
								tmpStr = tmpStr.replaceFirst("var[\\s]*=[\\s]*", "");									
								rowPlaceholder = tmpStr;

							} else {								
							}

						}	
						
						//解析colOfSerialNumber
						regEx = "colOfSerialNumber[\\s]*=[\\s]*[\"“”][\\s]*[0-9]*[\"“”]";
						pattern = Pattern.compile(regEx, Pattern.CASE_INSENSITIVE);
						m = pattern.matcher(attributeText);
						if (m.find()) {
							//String[] tmpArr = m.group(0).split("=");
							//String colOfSerialNumberStr = tmpArr[1].trim().replaceAll("[\"“”\\s　 ]*", "");
							String tmpStr = m.group(0);
							tmpStr = tmpStr.replaceFirst("colOfSerialNumber[\\s]*=[\\s]*[\"“”]", "");							
							String colOfSerialNumberStr = tmpStr.replaceFirst("[\"“”]$", "");		
							if (colOfSerialNumberStr != null && !colOfSerialNumberStr.trim().equals(""))
							{
								colOfSerialNumber = Integer.valueOf(colOfSerialNumberStr.trim());
							}

						} else {
							regEx = "colOfSerialNumber[\\s]*=[\\s]*[0-9]*[\\s]*";
							pattern = Pattern.compile(regEx, Pattern.CASE_INSENSITIVE);
							m = pattern.matcher(attributeText);
							if (m.find()) {
								//String[] tmpArr = m.group(0).split("=");
								//String colOfSerialNumberStr = tmpArr[1].trim().replaceAll("[\"“”\\s　 ]*", "");
								String tmpStr = m.group(0);
								tmpStr = tmpStr.replaceFirst("colOfSerialNumber[\\s]*=[\\s]*", "");									
								String colOfSerialNumberStr = tmpStr;
								if (colOfSerialNumberStr != null && !colOfSerialNumberStr.trim().equals(""))
								{
									colOfSerialNumber = Integer.valueOf(colOfSerialNumberStr.trim());
								}

							} else {								
							}

						}	
						
						//解析dateFormat
						regEx = "dateFormat[\\s]*=[\\s]*[\"“”][\\s]*[^\"“”]*[\"“”]";
						pattern = Pattern.compile(regEx, Pattern.CASE_INSENSITIVE);
						m = pattern.matcher(attributeText);
						if (m.find()) {
							//String[] tmpArr = m.group(0).split("=");
							//dateFormat = tmpArr[1].trim().replaceAll("[\"“”\\s　 ]*", "");
							String tmpStr = m.group(0);
							tmpStr = tmpStr.replaceFirst("dateFormat[\\s]*=[\\s]*[\"“”]", "");							
							dateFormat = tmpStr.replaceFirst("[\"“”]$", "");

						} else {
							regEx = "dateFormat[\\s]*=[\\s]*[^\\s\"“”]*[\\s]*";
							pattern = Pattern.compile(regEx, Pattern.CASE_INSENSITIVE);
							m = pattern.matcher(attributeText);
							if (m.find()) {
								//String[] tmpArr = m.group(0).split("=");
								//dateFormat = tmpArr[1].trim().replaceAll("[\"“”\\s　 ]*", "");
								String tmpStr = m.group(0);
								tmpStr = tmpStr.replaceFirst("dateFormat[\\s]*=[\\s]*", "");									
								dateFormat = tmpStr;

							} else {								
							}

						}	
						
						//解析datetimeFormat
						regEx = "datetimeFormat[\\s]*=[\\s]*[\"“”][\\s]*[^\"“”]*[\"“”]";
						pattern = Pattern.compile(regEx, Pattern.CASE_INSENSITIVE);
						m = pattern.matcher(attributeText);
						if (m.find()) {
							//String[] tmpArr = m.group(0).split("=");
							//datetimeFormat = tmpArr[1].trim().replaceAll("[\"“”\\s　 ]*", "");
							String tmpStr = m.group(0);
							tmpStr = tmpStr.replaceFirst("datetimeFormat[\\s]*=[\\s]*[\"“”]", "");							
							datetimeFormat = tmpStr.replaceFirst("[\"“”]$", "");

						} else {
							regEx = "datetimeFormat[\\s]*=[\\s]*[^\\s\"“”]*[\\s]*";
							pattern = Pattern.compile(regEx, Pattern.CASE_INSENSITIVE);
							m = pattern.matcher(attributeText);
							if (m.find()) {
								//String[] tmpArr = m.group(0).split("=");
								//datetimeFormat = tmpArr[1].trim().replaceAll("[\"“”\\s　 ]*", "");
								String tmpStr = m.group(0);
								tmpStr = tmpStr.replaceFirst("datetimeFormat[\\s]*=[\\s]*", "");									
								datetimeFormat = tmpStr;

							} else {								
							}

						}	
						
						//解析fieldsOfDatetime
						regEx = "fieldsOfDatetime[\\s]*=[\\s]*[\"“”][\\s]*[^\"“”]*[\"“”]";
						pattern = Pattern.compile(regEx, Pattern.CASE_INSENSITIVE);
						m = pattern.matcher(attributeText);
						if (m.find()) {
							//String[] tmpArr = m.group(0).split("=");
							//fieldsOfDatetime = tmpArr[1].trim().replaceAll("[\"“”\\s　 ]*", "");
							String tmpStr = m.group(0);
							tmpStr = tmpStr.replaceFirst("fieldsOfDatetime[\\s]*=[\\s]*[\"“”]", "");							
							fieldsOfDatetime = tmpStr.replaceFirst("[\"“”]$", "");

						} else {
							regEx = "fieldsOfDatetime[\\s]*=[\\s]*[^\\s\"“”]*[\\s]*";
							pattern = Pattern.compile(regEx, Pattern.CASE_INSENSITIVE);
							m = pattern.matcher(attributeText);
							if (m.find()) {
								//String[] tmpArr = m.group(0).split("=");
								//fieldsOfDatetime = tmpArr[1].trim().replaceAll("[\"“”\\s　 ]*", "");
								String tmpStr = m.group(0);
								tmpStr = tmpStr.replaceFirst("fieldsOfDatetime[\\s]*=[\\s]*", "");									
								fieldsOfDatetime = tmpStr;

							} else {								
							}

						}	
						
						//解析fieldsOfMergerow
						regEx = "fieldsOfMergerow[\\s]*=[\\s]*[\"“”][\\s]*[^\"“”]*[\"“”]";
						pattern = Pattern.compile(regEx, Pattern.CASE_INSENSITIVE);
						m = pattern.matcher(attributeText);
						if (m.find()) {
							//String[] tmpArr = m.group(0).split("=");
							//fieldsOfMergerow = tmpArr[1].trim().replaceAll("[\"“”\\s　 ]*", "");
							String tmpStr = m.group(0);
							tmpStr = tmpStr.replaceFirst("fieldsOfMergerow[\\s]*=[\\s]*[\"“”]", "");							
							fieldsOfMergerow = tmpStr.replaceFirst("[\"“”]$", "");

						} else {
							regEx = "fieldsOfMergerow[\\s]*=[\\s]*[^\\s\"“”]*[\\s]*";
							pattern = Pattern.compile(regEx, Pattern.CASE_INSENSITIVE);
							m = pattern.matcher(attributeText);
							if (m.find()) {
								//String[] tmpArr = m.group(0).split("=");
								//fieldsOfMergerow = tmpArr[1].trim().replaceAll("[\"“”\\s　 ]*", "");
								String tmpStr = m.group(0);
								tmpStr = tmpStr.replaceFirst("fieldsOfMergerow[\\s]*=[\\s]*", "");									
								fieldsOfMergerow = tmpStr;

							} else {								
							}

						}	
						
						//解析isomer
						regEx = "isomer[\\s]*=[\\s]*[\"“”][\\s]*\\[[\\s]*[\\s]*\\{[\\s]*[^\"“”]*\\}[\\s]*\\][\\s]*[\"“”]";
						pattern = Pattern.compile(regEx, Pattern.CASE_INSENSITIVE);
						m = pattern.matcher(attributeText);
						if (m.find()) {
							String tmpStr = m.group(0);
							tmpStr = tmpStr.replaceFirst("isomer[\\s]*=[\\s]*[\"“”]", "");							
							isomer = tmpStr.replaceFirst("[\"“”]$", "");

						} else {
							regEx = "isomer[\\s]*=[\\s]*\\[[\\s]*[\\s]*\\{[\\s]*[^\"“”]*\\}[\\s]*\\][\\s]*";
							pattern = Pattern.compile(regEx, Pattern.CASE_INSENSITIVE);
							m = pattern.matcher(attributeText);
							if (m.find()) {
								String tmpStr = m.group(0);
								tmpStr = tmpStr.replaceFirst("isomer[\\s]*=[\\s]*", "");									
								isomer = tmpStr;

							} else {								
							}

						}	
						
						
						
						
						
						//获取报表行模板
						if (beginIndex < endIndex)
						{
							for (int i=beginIndex + 1; i<endIndex; i++)
							{
								XWPFTableRow row = table.xwpfTable.getRow(i);
								List<Column> columnList = new ArrayList<Column>();
								if (row != null && row.getTableCells() != null && row.getTableCells().size() > 0)
								{
									for (int i4=0; i4<row.getTableCells().size(); i4++)
									{
										Column column = new Column();
										column.text = row.getCell(i4).getText();
										column.xwpfTableList = row.getCell(i4).getTables();
										column.index = i4;
										
										columnList.add(column);										
									}
								}
								
								TemplateRow templateRow = new TemplateRow();
								templateRow.tmpRow = row;
								templateRow.columnList = columnList;
								templateRow.index = i;
								
								rptRowTmp.add(templateRow);
							}
							
							result = new Section();
							result.table = table; //所属的msword表 
							result.lever = lever; //层深
							result.beginIndex = beginIndex; //开始索引
							result.endIndex = endIndex; //结束索引
							result.attributeText = attributeText; //属性文本
							result.alt = alt; //当数据源无数据时,用该文本替换
							result.colTitleIndex = colTitleIndex; //列标题行索引
							result.sql = sql; //取数据的sql语句
							result.dataSourceKey = dataSourceKey; //数据源的名称
							result.rowPlaceholder = rowPlaceholder; //数据行占位符
							result.colOfSerialNumber = colOfSerialNumber; //序号所在的列索引 索引以0开始 第一列为0
							result.dateFormat = dateFormat; //日期格式字符串 需要符合java.text.SampleDateFormat规范
							result.datetimeFormat = datetimeFormat;  //日期时间格式字符串 需要符合java.text.SimpleDateFormat规范
							result.fieldsOfDatetime = fieldsOfDatetime; //日期时间字段名称序列,多个已英文逗号分开， 形如    deptname,username 
							result.fieldsOfMergerow = fieldsOfMergerow; //需要合并行数据的列字段名称序列,多个已英文逗号分开， 形如    deptname,username 
							
							try {
								result.isomer = com.alibaba.fastjson.JSONArray.parseArray(isomer); //异构样式配置								
							} catch (Exception e) {
								result.isomer = null;
							}
							
							
							result.rptRowTmp = rptRowTmp; //行模板		

							
						}
						else
						{
							
						}
					
					}
					
					
				}				
			}
			
		}			
		
		return result;
		
	}
	
	/**
	 * 删除msword某一行
	 * @param section 要删除行所在的数据区域 
	 * @param removeIndex 要删除行的索引
	 * @param sectionList 数据区域列表
	 */
	public static void removeRow(Section section, int removeIndex, List<Section> sectionList) {
		if (section != null 
				&& section.table != null 
				&& section.table.xwpfTable != null 
				&& removeIndex > -1)
		{
			if ((section.colTitleIndex == removeIndex) || (section.beginIndex <= removeIndex && removeIndex <= section.endIndex))
			{//验证removeIndex在数据区域 section中
				section.table.xwpfTable.removeRow(removeIndex);
				if (section.colTitleIndex == removeIndex)
				{
					section.colTitleIndex = -1;		
					section.beginIndex --;
					if (section.rptRowTmp != null && section.rptRowTmp.size() > 0)
					{
						for (int i=0; i<section.rptRowTmp.size(); i++)
						{
							TemplateRow templateRow = section.rptRowTmp.get(i);	
							if (templateRow != null)
							{
								templateRow.index --;
							}
							
						}
					}					
					section.endIndex --;
				}
				else if (section.beginIndex == removeIndex)
				{
					section.beginIndex = -1;		
					if (section.rptRowTmp != null && section.rptRowTmp.size() > 0)
					{
						for (int i=0; i<section.rptRowTmp.size(); i++)
						{
							TemplateRow templateRow = section.rptRowTmp.get(i);	
							if (templateRow != null)
							{
								templateRow.index --;
							}
							
						}
					}
					section.endIndex --;
				}
				else if (section.endIndex == removeIndex)
				{
					section.endIndex = -1;				
				}	
				else if (section.rptRowTmp != null && section.rptRowTmp.size() > 0)
				{
					for (int i=0; i<section.rptRowTmp.size(); i++)
					{
						TemplateRow templateRow = section.rptRowTmp.get(i);	
						if (templateRow != null && templateRow.index == removeIndex)
						{
							section.rptRowTmp.remove(templateRow);
							for (int i1=i+1; i1<section.rptRowTmp.size(); i1++)
							{
								TemplateRow templateRow1 = section.rptRowTmp.get(i1);	
								if (templateRow1 != null)
								{
									templateRow1.index --;
								}
							}
							section.endIndex --;
							break;
						}
						
					}
				}			
				for (int j=section.selfIndex+1; j<section.table.sectionList.size(); j++)
				{
					Section tar = section.table.sectionList.get(j);
					if (tar != null && tar.rptRowTmp != null && tar.rptRowTmp.size() > 0)
					{
						tar.beginIndex --;
						tar.endIndex --;				
						
						for (int i2=0; i2<tar.rptRowTmp.size(); i2++)
						{
							TemplateRow templateRow = tar.rptRowTmp.get(i2);	
							if (templateRow != null)
							{
								templateRow.index --;
							}
							
						}
					

					}
					
				}
				
				
			}
			else
			{//不处理
				
			}
			
		}
	}
	
	/**
	 * 合并某列行数据
	 * @param table msword表格
	 * @param col 列索引
	 * @param fromRow 合并开始那个行的索引
	 * @param toRow 合并结束那个行的索引
	 */
	public static void mergeCellVertically(XWPFTable table, int col, int fromRow, int toRow) {
        for (int rowIndex = fromRow; rowIndex <= toRow; rowIndex++)
        {
            CTVMerge vmerge = CTVMerge.Factory.newInstance();
            if(rowIndex == fromRow)
            {
                // The first merged cell is set with RESTART merge value
                vmerge.setVal(STMerge.RESTART);
            } 
            else 
            {
                // Cells which join (merge) the first one, are set with CONTINUE
                vmerge.setVal(STMerge.CONTINUE);
            }
            XWPFTableCell cell = table.getRow(rowIndex).getCell(col);
            // Try getting the TcPr. Not simply setting an new one every time.
            CTTcPr tcPr = cell.getCTTc().getTcPr();
            if (tcPr != null) 
            {
                tcPr.setVMerge(vmerge);
            } 
            else 
            {
                // only set an new TcPr if there is not one already
                tcPr = CTTcPr.Factory.newInstance();
                tcPr.setVMerge(vmerge);
                cell.getCTTc().setTcPr(tcPr);
            }
        }
    }

	/**
	 * 合并某行列数据
	 * @param table  msword表格
	 * @param row 行索引
	 * @param fromCol  合并开始那个列的索引
	 * @param toCol  合并结束那个列的索引
	 */
	public static void mergeCellHorizontally(XWPFTable table, int row, int fromCol, int toCol) {
        for(int colIndex = fromCol; colIndex <= toCol; colIndex++)
        {
            CTHMerge hmerge = CTHMerge.Factory.newInstance();
            if(colIndex == fromCol)
            {
                // The first merged cell is set with RESTART merge value
                hmerge.setVal(STMerge.RESTART);
            } 
            else 
            {
                // Cells which join (merge) the first one, are set with CONTINUE
                hmerge.setVal(STMerge.CONTINUE);
            }
            XWPFTableCell cell = table.getRow(row).getCell(colIndex);
            // Try getting the TcPr. Not simply setting an new one every time.
            CTTcPr tcPr = cell.getCTTc().getTcPr();
            if (tcPr != null) 
            {
                tcPr.setHMerge(hmerge);
            } 
            else 
            {
                // only set an new TcPr if there is not one already
                tcPr = CTTcPr.Factory.newInstance();
                tcPr.setHMerge(hmerge);
                cell.getCTTc().setTcPr(tcPr);
            }
        }
    }
	
	
	/**
	 * 递归处理数据行样式
	 * @param section
	 
	public static void weaveStyleTree(Section section) {
		
		if (section != null)
		{
			weaveStyle(section, section.table.report.varValues);
			
		}
	}*/
	
	/**
	 * 处理数据行样式
	 * @param section
	 * @return
	 
	public static Section weaveStyle(Section section, Map<String, Object> varValues) {
		
		if (section != null 
				&& varValues != null
				&& section.table != null 
				&& section.table.report != null 				
				&& section.dataSourceKey != null 
				&& !section.dataSourceKey.trim().equals("")
				&& section.rowPlaceholder != null 
				&& !section.rowPlaceholder.trim().equals(""))
		{
			List<List<Integer>> appendRptRows = new ArrayList<List<Integer>>(); //根据数据源的记录数添加的表格行索引列表

			if (varValues.isEmpty())
			{									
			}
			else
			{
				if (section.rptRowTmp != null && section.rptRowTmp.size() > 0)
				{										
					Object dataSource = varValues.get(section.dataSourceKey);
					if (dataSource != null)
					{
						if (dataSource instanceof List)
						{
							List items = (List)dataSource;
							if (items.size() > 0)
							{
								//按照报表行模板添加报表行
								for (int i1=0; i1<items.size(); i1++)
								{
									Object rowSource = items.get(i1);
									if (rowSource != null)
									{
										if (rowSource instanceof Map)
										{
											Map item = (Map)rowSource;
											
											List<Integer> rptRow = new ArrayList<Integer>();
											
											for (int i2=0; i2<section.rptRowTmp.size(); i2++)
											{
												XWPFTableRow row = section.rptRowTmp.get(i2);							
												
												boolean issucc = section.table.xwpfTable.addRow(row, section.endIndex);
												if (issucc)
												{
													rptRow.add(section.endIndex);
													
													section.endIndex ++;
												}
												
											}
											
											appendRptRows.add(rptRow);
											
											
											
											msword.report.Operate.saveas(section.table.report,"D:\\TEMP\\OOOO\\javamsword\\result101"+section.dataSourceKey+".docx");
											
											
											if (section.children != null && section.children.size() > 0)
											{												
												item.put("children", copyTree(section.children));
												
												for (int i=0; i<section.children.size(); i++)
												{
													Section tar = section.children.get(i);
													
													weaveStyle(tar, item);


													
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
		
			section.appendRptRows = appendRptRows;
			
			if (appendRptRows.size() > 0 
					&& section.table.sectionList != null 
					&& section.table.sectionList.size() > 0)
			{				
				for (int j=section.selfIndex+1; j<section.table.sectionList.size(); j++)
				{
					Section tar = section.table.sectionList.get(j);
					if (tar != null)
					{
						tar.beginIndex += appendRptRows.size();
						tar.endIndex += appendRptRows.size();
					}
					
				}
			}
			
		
		}
		
		return section;
	}*/

	
	

}
