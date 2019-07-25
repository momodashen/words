package msword;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.script.ScriptEngine;
import javax.script.ScriptEngineManager;
import javax.script.ScriptException;

import msword.report.Picture;

import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public class Operate {
	
	/**
	 * 解析输入文本
	 * @param placeholder  数据占位符
	 * @param dateFormat 日期格式字符串 需要符合java.text.SampleDateFormat规范
	 * @param datetimeFormat 日期时间格式字符串 需要符合java.text.SimpleDateFormat规范
	 * @param fieldsOfDatetime 日期时间字段名称序列,多个已英文逗号分开， 形如    deptname,username 
	 * @param isomer 异构样式配置
	 * @param text 输入文本
	 * @param varValues 存放变量占位符及对应值的map
	 * @return
	 */
	public static String textParse(String placeholder, String dateFormat, String datetimeFormat, String fieldsOfDatetime, com.alibaba.fastjson.JSONArray isomer, String text, Map<String, Object> varValues) {
		String result = text;
		
		Map<String, List<Map<String, Object>>> tmp = textParseToRuns(text, null, null);
		List<Map<String, Object>> allList = tmp.get("allList");
		List<Map<String, Object>> varList = tmp.get("varList");
				
		result = varAssign(placeholder, dateFormat, datetimeFormat, fieldsOfDatetime, isomer, allList, varList, varValues, null);	
		
		
		return result;
	}
	
	/**
	 * 解析文本的在msword中的run列表
	 * @param placeholder  数据占位符
	 * @param dateFormat 日期格式字符串 需要符合java.text.SampleDateFormat规范
	 * @param datetimeFormat 日期时间格式字符串 需要符合java.text.SimpleDateFormat规范
	 * @param fieldsOfDatetime 日期时间字段名称序列,多个已英文逗号分开， 形如    deptname,username 
	 * @param isomer 异构样式配置
	 * @param text 输入文本
	 * @param runs 输入文本的在msword中的run列表
	 * @param varValues 存放变量占位符及对应值的map
	 * @return
	 */
	public static String runsParse(String placeholder, String dateFormat, String datetimeFormat, String fieldsOfDatetime, com.alibaba.fastjson.JSONArray isomer, String text, List<XWPFRun> runs, Map<String, Object> varValues) {
		String result = text;		
		
		Map<String, List<Map<String, Object>>> tmp = textParseToRuns(text, null, null);
		List<Map<String, Object>> allList = tmp.get("allList");
		List<Map<String, Object>> varList = tmp.get("varList");
		
		entanglement(text, runs, allList, varList);
				
		result = varAssign(placeholder, dateFormat, datetimeFormat, fieldsOfDatetime, isomer, allList, varList, varValues, runs);	
		
		
		return result;
	}	
	
	/**
	 * 给一般字符串和变量列表中的变量项占位符替换成变量值返回字符串
	 * @param placeholder  数据占位符
	 * @param dateFormat 日期格式字符串 需要符合java.text.SampleDateFormat规范
	 * @param datetimeFormat 日期时间格式字符串 需要符合java.text.SimpleDateFormat规范
	 * @param fieldsOfDatetime 日期时间字段名称序列,多个已英文逗号分开， 形如    deptname,username 
	 * @param isomer 异构样式配置
	 * @param allList  存放一般字符串和变量的列表
	 * @param varList 存放变量的列表
	 * @param varValues 存放变量占位符及对应值的map
	 * @param runs 文本的在msword中的run列表
	 * @return
	 */
	public static String varAssign(String placeholder, String dateFormat, String datetimeFormat, String fieldsOfDatetime, com.alibaba.fastjson.JSONArray isomer, List<Map<String, Object>> allList, List<Map<String, Object>> varList, Map<String, Object> varValues, List<XWPFRun> runs) {
		String result = "";
		
		allList = varAssignForList(placeholder, dateFormat, datetimeFormat, fieldsOfDatetime, isomer, allList, varList, varValues, runs);
		
		if (allList != null && allList.size() > 0)
		{			
			for (int i=0; i<allList.size(); i++)
			{
				Map<String, Object> one =  allList.get(i);
				String content = (String) one.get("content");
				if (content != null)
				{
					result += content;
				}				
			}
					
		}
		
		return result;
	}
	
	/**
	 * 给一般字符串和变量列表中的变量项占位符替换成变量值
	 * @param placeholder  数据占位符
	 * @param dateFormat 日期格式字符串 需要符合java.text.SampleDateFormat规范
	 * @param datetimeFormat 日期时间格式字符串 需要符合java.text.SimpleDateFormat规范
	 * @param fieldsOfDatetime 日期时间字段名称序列,多个已英文逗号分开， 形如    deptname,username 
	 * @param isomer 异构样式配置
	 * @param allList  存放一般字符串和变量的列表
	 * @param varList 存放变量的列表
	 * @param varValues 存放变量占位符及对应值的map
	 * @param runs 文本的在msword中的run列表
	 * @return
	 */
	public static List<Map<String, Object>> varAssignForList(String placeholder, String dateFormat, String datetimeFormat, String fieldsOfDatetime, com.alibaba.fastjson.JSONArray isomer, List<Map<String, Object>> allList, List<Map<String, Object>> varList, Map<String, Object> varValues, List<XWPFRun> runs) {
		List<Map<String, Object>> result = allList;
		if (allList != null && allList.size() > 0)
		{
			if (varList != null && varList.size() > 0)
			{				
				if (runs != null && runs.size() > 0)
				{
					List runSectList = runsSubsection(varList);
					if (runSectList != null && runSectList.size() > 0)
					{
						for (int j=0; j<runSectList.size(); j++)
						{
							Map runSect = (Map) runSectList.get(j);
							if (runSect != null)
							{
								String tarRunIndex = (String) runSect.get("runIndex");
								List tarVarMap = (List) runSect.get("varMap");
								
								if (tarRunIndex != null 
										&& !tarRunIndex.trim().equals("") 
										&& tarVarMap != null 
										&& tarVarMap.size() > 0)
								{
									String[] runIndexArr = tarRunIndex.trim().split(",");
									if (runIndexArr != null && runIndexArr.length > 0)
									{
										String text = "";
										for (int j1=0; j1<runIndexArr.length; j1++)
										{
											String runIndex = runIndexArr[j1];
											if (runIndex != null && !runIndex.trim().equals(""))
											{											
												XWPFRun run = runs.get(Integer.valueOf(runIndex));
												text += run.getText(0);
											}
										}
										if (text != null && !text.trim().equals("")) 
										{
											com.alibaba.fastjson.JSONObject style = isomerStyle(text, varValues, placeholder, dateFormat, datetimeFormat, fieldsOfDatetime, isomer);
											
											boolean isSetText = false;
											if (varValues != null)
											{
												Map tar = textAnaly(varValues, 0L, placeholder, dateFormat, datetimeFormat, fieldsOfDatetime, text);
												if (tar != null)
												{
													isSetText = (Boolean)tar.get("isSetText");
													text = (String)tar.get("text");
												}
											}
											else
											{
												for (int i=0; i<varList.size(); i++)
												{
													Map<String, Object> var =  varList.get(i);
													if (var != null)
													{
														String content = (String) var.get("content");
														if (content != null && !content.trim().equals(""))
														{
															isSetText = true;
															text = text.replace(content, "");
														}
														
													}
													
												}							
											}
											if (isSetText)
											{
												int runCounter = 0;
												for (int j1=0; j1<runIndexArr.length; j1++)
												{
													String runIndex = runIndexArr[j1];
													if (runIndex != null && !runIndex.trim().equals(""))
													{	
														XWPFRun run = runs.get(Integer.valueOf(runIndex));
														
														if (style != null)
														{
															String color = style.getString("color");
															if (color != null && !color.trim().equals(""))
															{
																run.setColor(color);
															}
														}
														
														runCounter ++;
														if (runCounter == 1)
														{//解决回车问题
															if (text.contains("\n"))
															{
																String[] arr = text.split("\n");
																if (arr != null && arr.length > 0)
																{
																	for (int y=0; y<arr.length; y++)
																	{
																		if (y == 0)
																		{
																			run.setText(arr[y], 0);
																		}
																		else
																		{
																			run.setText(arr[y]);
																		}
//																		run.addCarriageReturn();		
																		run.addBreak();
																		
																		
																	}
																}
															}		
															else
															{
																run.setText(text, 0);
															}																												
														}
														else
														{
															run.setText("", 0);
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
					
/*					
					int startRunIndex = runs.size() - 1;
					int stopRunIndex = -1;
					for (int i=0; i<varList.size(); i++)
					{
						Map<String, Object> var =  varList.get(i);
						if (var != null && var.get("content") != null && var.get("indexOfAllList") != null)
						{
							String indexOfRunsStr = (String) var.get("indexOfRunsStr");
							if (indexOfRunsStr != null && !indexOfRunsStr.trim().equals(""))
							{
								String[] indexOfRunArr = indexOfRunsStr.trim().split(",");
								if (indexOfRunArr != null && indexOfRunArr.length > 0)
								{
									for (int i1=0; i1<indexOfRunArr.length; i1++)
									{
										if (indexOfRunArr[i1] != null && !indexOfRunArr[i1].trim().equals(""))
										{
											int indexOfRun = Integer.valueOf(indexOfRunArr[i1]);										
											if (indexOfRun > stopRunIndex)
											{
												stopRunIndex = indexOfRun;
											}
											if (indexOfRun < startRunIndex)
											{
												startRunIndex = indexOfRun;
											}
										}
									}
								}
							}
						}
					}
					
					if (stopRunIndex > -1)
					{

						String text = "";
						for (int i=startRunIndex; i<=stopRunIndex; i++)
						{
							XWPFRun run = runs.get(i);
							text += run.getText(0);
						}
						if (text != null && !text.trim().equals("")) 
						{
							boolean isSetText = false;
							if (varValues != null)
							{
								for (Entry<String, Object> entry : varValues.entrySet()) 
								{
									String key = entry.getKey();								
									String regEx = "\\$\\{[\\s　  ]*" + key + "[\\s　  ]*\\}";
									if (placeholder != null && !placeholder.trim().equals(""))
									{
										regEx = "\\$\\{[\\s　  ]*" + placeholder + "\\.[\\s　  ]*" + key + "[\\s　  ]*\\}";
									}
									
									Pattern pattern = Pattern.compile(regEx, Pattern.CASE_INSENSITIVE);
									Matcher m = pattern.matcher(text);
									if (m.find()) {
										isSetText = true;
										Object value = entry.getValue();
										if (value != null)
										{
											if (value instanceof java.util.Date) 
											{												
												// 日期格式化替换
												if (fieldsOfDatetime != null && !fieldsOfDatetime.trim().equals(""))
												{
													String[] fieldsOf = fieldsOfDatetime.trim().split(",");
													if (fieldsOf != null && fieldsOf.length > 0) { } else
													{
														fieldsOf = fieldsOfDatetime.trim().split("，");
													}
													if (fieldsOf != null && fieldsOf.length > 0) { } else
													{
														fieldsOf = fieldsOfDatetime.trim().split(";");
													}
													if (fieldsOf != null && fieldsOf.length > 0) { } else
													{
														fieldsOf = fieldsOfDatetime.trim().split("；");
													}
													
													boolean isDatetime = false;
													for (int i4=0; i4<fieldsOf.length; i4++)
													{
														String fieldOf = fieldsOf[i4];
														if (fieldOf != null && !fieldOf.trim().equals(""))
														{
															if (fieldOf.trim().equalsIgnoreCase(key.trim()))
															{
																isDatetime = true;
															}															
														}														
													}																										
													if (isDatetime)
													{//作时间格式
														if (datetimeFormat != null && !datetimeFormat.trim().equals(""))
														{
														}
														else
														{
															datetimeFormat = "yyyy年MM月dd日 HH时mm分ss秒";
														}
														text = text.replaceAll(regEx, (new java.text.SimpleDateFormat(datetimeFormat)).format((java.util.Date)value));
													}
													else
													{//作日期格式
														if (dateFormat != null && !dateFormat.trim().equals(""))
														{
														}
														else
														{
															dateFormat = "yyyy年MM月dd日";
														}
														text = text.replaceAll(regEx, (new java.text.SimpleDateFormat(dateFormat)).format((java.util.Date)value));
														
													}
												}
												else
												{//作日期格式
													if (dateFormat != null && !dateFormat.trim().equals(""))
													{
													}
													else
													{
														dateFormat = "yyyy年MM月dd日";
													}
													text = text.replaceAll(regEx, (new java.text.SimpleDateFormat(dateFormat)).format((java.util.Date)value));
													
												}
												
												System.out.println(text);
											
											}
											else// if (value instanceof String) 
											{										
												// 文本替换
												text = text.replaceAll(regEx, value.toString());
												
												System.out.println(text);
											}
										}
										else
										{
											//空字符处理										
											text = text.replaceAll(regEx, "");
										}
										
									}
									
									
								
								}
							}
							else
							{
								for (int i=0; i<varList.size(); i++)
								{
									Map<String, Object> var =  varList.get(i);
									if (var != null)
									{
										String content = (String) var.get("content");
										if (content != null && !content.trim().equals(""))
										{
											isSetText = true;
											text = text.replace(content, "");
										}
										
									}
									
								}							
							}
							if (isSetText) 
							{
								int runCounter = 0;
								for (int i=startRunIndex; i<=stopRunIndex; i++)
								{
									XWPFRun run = runs.get(i);
									
									runCounter ++;
									if (runCounter == 1)
									{//解决回车问题
										if (text.contains("\n"))
										{
											String[] arr = text.split("\n");
											if (arr != null && arr.length > 0)
											{
												for (int y=0; y<arr.length; y++)
												{
													if (y == 0)
													{
														run.setText(arr[y], 0);
													}
													else
													{
														run.setText(arr[y]);
													}
//													run.addCarriageReturn();		
													run.addBreak();
													
												}
											}
										}		
										else
										{
											run.setText(text, 0);
										}														
										
									}
									else
									{
										run.setText("", 0);
									}													
								}
							}
						}					
					}
*/					
					
				}
				
				for (int i=0; i<varList.size(); i++)
				{
					Map<String, Object> var =  varList.get(i);
					if (var != null && var.get("content") != null && var.get("indexOfAllList") != null)
					{
						String content = (String) var.get("content");
						Integer indexOfAllList = (Integer) var.get("indexOfAllList");
						String indexOfRunsStr = (String) var.get("indexOfRunsStr");
						
						String vartext = content;
						if (vartext != null) 
						{
							boolean isSetText = false;
							if (varValues != null)
							{
								Map tar = textAnaly(varValues, 1L, placeholder, dateFormat, datetimeFormat, fieldsOfDatetime, vartext);
								if (tar != null)
								{
									isSetText = (Boolean)tar.get("isSetText");
									vartext = (String)tar.get("text");
								}
								
/*								
								for (Entry<String, Object> entry : varValues.entrySet()) 
								{
									String key = entry.getKey();								
									String regEx = "\\$\\{[\\s　  ]*" + key + "[\\s　  ]*\\}";
									if (placeholder != null && !placeholder.trim().equals(""))
									{
										regEx = "\\$\\{[\\s　  ]*" + placeholder + "\\.[\\s　  ]*" + key + "[\\s　  ]*\\}";
									}
									
									Pattern pattern = Pattern.compile(regEx, Pattern.CASE_INSENSITIVE);
									Matcher m = pattern.matcher(vartext);
									if (m.find()) {
										isSetText = true;
										Object value = entry.getValue();
										if (value != null)
										{
											if (value instanceof java.util.Date) 
											{												
												// 日期格式化替换
												if (fieldsOfDatetime != null && !fieldsOfDatetime.trim().equals(""))
												{
													String[] fieldsOf = fieldsOfDatetime.trim().split(",");
													if (fieldsOf != null && fieldsOf.length > 0) { } else
													{
														fieldsOf = fieldsOfDatetime.trim().split("，");
													}
													if (fieldsOf != null && fieldsOf.length > 0) { } else
													{
														fieldsOf = fieldsOfDatetime.trim().split(";");
													}
													if (fieldsOf != null && fieldsOf.length > 0) { } else
													{
														fieldsOf = fieldsOfDatetime.trim().split("；");
													}
													
													boolean isDatetime = false;
													for (int i4=0; i4<fieldsOf.length; i4++)
													{
														String fieldOf = fieldsOf[i4];
														if (fieldOf != null && !fieldOf.trim().equals(""))
														{
															if (fieldOf.trim().equalsIgnoreCase(key.trim()))
															{
																isDatetime = true;
															}															
														}														
													}																										
													if (isDatetime)
													{//作时间格式
														if (datetimeFormat != null && !datetimeFormat.trim().equals(""))
														{
														}
														else
														{
															datetimeFormat = "yyyy年MM月dd日 HH时mm分ss秒";
														}
														vartext = vartext.replaceAll(regEx, (new java.text.SimpleDateFormat(datetimeFormat)).format((java.util.Date)value));
													}
													else
													{//作日期格式
														if (dateFormat != null && !dateFormat.trim().equals(""))
														{
														}
														else
														{
															dateFormat = "yyyy年MM月dd日";
														}
														vartext = vartext.replaceAll(regEx, (new java.text.SimpleDateFormat(dateFormat)).format((java.util.Date)value));
														
													}
												}
												else
												{//作日期格式
													if (dateFormat != null && !dateFormat.trim().equals(""))
													{
													}
													else
													{
														dateFormat = "yyyy年MM月dd日";
													}
													vartext = vartext.replaceAll(regEx, (new java.text.SimpleDateFormat(dateFormat)).format((java.util.Date)value));
													
												}
												
												System.out.println(vartext);
											
											}
											else// if (value instanceof String) 
											{										
												// 文本替换
												vartext = vartext.replaceAll(regEx, value.toString());
												
												System.out.println(vartext);
											}
										}
										else
										{//空字符处理										
											vartext = "";
										}
									}
								}
*/								
								
							}
							else
							{
								isSetText = true;
								vartext = "";
							}
							if (isSetText) 
							{
								Map<String, Object> one = allList.get(indexOfAllList);
								one.put("content", vartext);
								
								allList.set(indexOfAllList, one);
								
							}
						}
						
					}
					
				}	
			}
					
		}
		
		return result;
	}
	
	/**
	 * 变量涉及单元格的XWPFRun序列分段
	 * @param varList 文本解析后的变量列表
	 * @return
	 */
	public static List<Map> runsSubsection(List<Map<String, Object>> varList) {
		List<Map> result = null;
		
		if (varList != null && varList.size() > 0)
		{
			result = new ArrayList<Map>();			
			
			for (int i=0; i<varList.size(); i++)
			{
				Map<String, Object> var =  varList.get(i);
				if (var != null && var.get("content") != null && var.get("indexOfAllList") != null)
				{
					String indexOfRunsStr = (String) var.get("indexOfRunsStr");
					if (indexOfRunsStr != null && !indexOfRunsStr.trim().equals(""))
					{
						String startIndexStr = indexOfRunsStr;
						int startIndexOf = indexOfRunsStr.indexOf(",");
						if (startIndexOf > -1)
						{
							startIndexStr = indexOfRunsStr.substring(0,startIndexOf);
						}
						int startIndex = Integer.valueOf(startIndexStr);						
						
						if (result.size() > 0)
						{
							Map iTar = result.get(result.size() - 1);
							String tarRunIndex = (String) iTar.get("runIndex");
							List tarVarMap = (List) iTar.get("varMap");
							
							String tmpEndIndexStr = tarRunIndex;
							int tmpEndIndexOf = tarRunIndex.lastIndexOf(",");
							if (tmpEndIndexOf > -1)
							{
								tmpEndIndexStr = tarRunIndex.substring(tmpEndIndexOf + 1);
							}
							int tmpEndIndex = Integer.valueOf(tmpEndIndexStr);
							
							if (tmpEndIndex >= startIndex)
							{//当前变量所在区域与上一个变量所在区域无缝一体
								
								tarRunIndex = stringSeqJoin(indexOfRunsStr, tarRunIndex);
								tarVarMap.add(var);
								
								iTar.put("runIndex", tarRunIndex);
								iTar.put("varMap", tarVarMap);
								
								result.set(result.size() - 1, iTar);
							}
							else
							{
								tarRunIndex = indexOfRunsStr;
								tarVarMap = new ArrayList();
								tarVarMap.add(var);
								
								iTar = new HashMap();
								iTar.put("runIndex", tarRunIndex);
								iTar.put("varMap", tarVarMap);
								
								result.add(iTar);
								
							}
							
						}
						else
						{
							String tarRunIndex = indexOfRunsStr;
							List tarVarMap = new ArrayList();
							tarVarMap.add(var);
							
							Map iTar = new HashMap();
							iTar.put("runIndex", tarRunIndex);
							iTar.put("varMap", tarVarMap);
							
							result.add(iTar);
						}
						
						
					}
				}
			}
		}		
		
		
		return result;
	}
	
	/**
	 * 不重复连接字符串序列
	 * @param text 
	 * @param target 目标字符串
	 * @return
	 */
	public static String stringSeqJoin(String text, String target) {
		String result = target;
		if (text != null && !text.trim().equals(""))
		{
			if (target != null && !target.trim().equals(""))
			{
				String[] txtArr = text.trim().split(",");
				if (txtArr != null && txtArr.length > 0)
				{
					for (int i=0; i<txtArr.length; i++)
					{
						String txt = txtArr[i];
						if (txt != null && !txt.trim().equals(""))
						{							
							if (("," + result.trim() + ",").indexOf("," + txt.trim() + ",") > -1) 
							{//已存在								
							}
							else
							{
								result += "," + txt.trim();
							}							
						}						
					}					
				}				
			}
			else
			{
				result = text;
			}
		}
		
		return result;
	}

	/**
	 * 解析关系
	 * @param text 文本
	 * @param runs 文本的在msword中的run列表
	 * @param allList 文本解析后的字符及变量列表
	 * @param varList 文本解析后的变量列表
	 * @return
	 */
	public static List<Map<String, Object>> entanglement(String text, List<XWPFRun> runs, List<Map<String, Object>> allList, List<Map<String, Object>> varList) {
		List<Map<String, Object>> result = null;
		
		if (text != null && !text.trim().equals("") && varList != null && varList.size() > 0)
		{			
			result = new ArrayList<Map<String, Object>>();			
			
			String tmp_text = "";
			char[] charArr = text.toCharArray();
			for (int i=0; i<charArr.length; i++)
			{				
				char content = charArr[i];		
				tmp_text += String.valueOf(content);				
				
				int indexOfRun = indexOfRunList(i, tmp_text, runs);
				int indexOfAll = indexOfAllList(i, tmp_text, allList);
				
				Map longlati = new HashMap<String, Object>();
				longlati.put("content", content);
				longlati.put("indexOfRun", indexOfRun);
				longlati.put("indexOfAll", indexOfAll);				
				
				result.add(longlati);
				
				if (indexOfAll > -1 && indexOfRun > -1)
				{
					Map<String, Object> one = allList.get(indexOfAll);
					if (one != null && one.get("type") != null && ((String)one.get("type")).equals("var"))
					{
						int indexOfVarList = (Integer) one.get("indexOfVarList");
						Map<String, Object> var = varList.get(indexOfVarList);
						
						String indexOfRunsStr = (String) var.get("indexOfRunsStr");
						if (indexOfRunsStr != null && !indexOfRunsStr.trim().equals(""))
						{
							if (("," + indexOfRunsStr.trim() + ",").endsWith("," + String.valueOf(indexOfRun) + ","))
							{//过滤重复								
							}
							else
							{
								indexOfRunsStr += "," + String.valueOf(indexOfRun);
							}													
						}
						else
						{
							indexOfRunsStr = String.valueOf(indexOfRun);
						}
						
						var.put("indexOfRunsStr", indexOfRunsStr);
						varList.set(indexOfVarList, var);
						
					}
				}
				
			}	
			
		}		
		
		return result;
	}	
	
	/**
	 * 解析输入文本为一般字符串和变量列表
	 * @param text 输入文本
	 * @param allList 存放一般字符串和变量的列表
	 * @param varList 存放变量的列表
	 * @return
	 */
	public static Map<String, List<Map<String, Object>>> textParseToRuns(String text, List<Map<String, Object>> allList, List<Map<String, Object>> varList) {
		
//		if (text != null && !text.trim().equals(""))
		if (text != null)
		{
			int index0 = text.indexOf("${");
			int index1 = text.indexOf("}");
			
			if (index0 > -1 && index1 > -1)
			{
				if (index1 > index0)
				{
					String text_tmp = text.substring(index0, index1 + 1); //“${”“}”之间的字符串
					int index_tmp = text_tmp.lastIndexOf("${");
					index0 = index0 + index_tmp;
					
					
					String text__1 = text.substring(0, index0); //“${”前面的字符串
					String text_0 = text.substring(index0, index1 + 1); //恰当的“${”“}”之间的字符串
					String text_1 = text.substring(index1 + 1, text.length()); //“}”后面的字符串
					
					Map<String, List<Map<String, Object>>> recursion = textParseToRuns(text__1, allList, varList);
					allList = recursion.get("allList");
					varList = recursion.get("varList");
					
					if (allList == null)
					{
						allList = new ArrayList<Map<String, Object>>();					
					}
					Map<String, Object> one = new HashMap<String, Object>();
					one.put("content", text_0);
					one.put("type", "var");		
					allList.add(one);
					Integer indexOfAllList = allList.size() - 1; //取得变量推入all列表中的索引
					if (varList == null)
					{
						varList = new ArrayList<Map<String, Object>>();					
					}					
					Map<String, Object> var = new HashMap<String, Object>();
					var.put("content", text_0);
					var.put("indexOfAllList", indexOfAllList);					
					varList.add(var); //将变量推入变量列表
					Integer indexOfVarList = varList.size() - 1; //取得变量推入var列表中的索引
					
					one.put("indexOfVarList", indexOfVarList);	
					allList.set(indexOfAllList, one);
					
							
					recursion = textParseToRuns(text_1, allList, varList);
					allList = recursion.get("allList");
					varList = recursion.get("varList");		
							
				}
				else
				{
					String text__1 = text.substring(0, index1 + 1); //“}”前面的字符串
					String text_1 = text.substring(index1 + 1, text.length()); //“}”后面的字符串
					if (allList == null)
					{
						allList = new ArrayList<Map<String, Object>>();					
					}
					Map<String, Object> one = new HashMap<String, Object>();
					one.put("content", text__1);
					one.put("type", "str");		
					allList.add(one);
					
					Map<String, List<Map<String, Object>>> recursion = textParseToRuns(text_1, allList, varList);
					allList = recursion.get("allList");
					varList = recursion.get("varList");
					
				}
				
			}
			else
			{
				if (allList == null)
				{
					allList = new ArrayList<Map<String, Object>>();					
				}
				Map<String, Object> one = new HashMap<String, Object>();
				one.put("content", text);
				one.put("type", "str");	
				allList.add(one);
			}			

		}
		
		Map<String, List<Map<String, Object>>> result = new HashMap<String, List<Map<String, Object>>>();
		result.put("allList", allList);
		result.put("varList", varList);
		
		
		return result;
	}
	
	/**
	 * 获取字符索引index在msword中的run列表中的索引
	 * @param index 字符索引
	 * @param tarText 目标字符
	 * @param runs 文本的在msword中的run列表
	 * @return
	 */
	public static int indexOfRunList(int index, String tarText, List<XWPFRun> runs) {
		int result = -1;
		
		if (runs != null && runs.size() > 0)
		{
			String tmp_text = "";
			int offset = 0;
			for (int i=0; i<runs.size(); i++)
			{
				if (runs.get(i) != null && runs.get(i).getText(0) != null)
				{
					tmp_text += runs.get(i).getText(0);
					offset += runs.get(i).getText(0).length();				
				}				
				
//				if (tmp_text.length() >= tarText.length())
				if (offset > index)
				{
					result = i;
					break;
				}
				
			}			
		}
		
		return result;
	}
	
	/**
	 * 获取字符索引index在解析字符及变量列表中的索引
	 * @param index 字符索引
	 * @param tarText 目标字符
	 * @param allList 解析字符及变量列表
	 * @return
	 */
	public static int indexOfAllList(int index, String tarText, List<Map<String, Object>> allList) {
		int result = -1;
		
		if (allList != null && allList.size() > 0)
		{
//			System.out.println(tarText);
			
			String tmp_text = "";
			int offset = 0;
			for (int i=0; i<allList.size(); i++)
			{
				Map<String, Object> one = allList.get(i);				
				tmp_text += (String)one.get("content");
//				System.out.println(tmp_text);				
				
				offset += ((String)one.get("content")).length();				
				
//				if (tmp_text.length() >= tarText.length())
				if (offset > index)
				{
					result = i;
					break;
				}
				
			}			
		}
		
		return result;
	}

	
	/**
	 * 异构样式配置解析出样式
	 * @param text 文本
	 * @param varValues 存放变量占位符及对应值的map
	 * @param placeholder 数据占位符
	 * @param dateFormat 日期格式字符串 需要符合java.text.SampleDateFormat规范
	 * @param datetimeFormat 日期时间格式字符串 需要符合java.text.SimpleDateFormat规范
	 * @param fieldsOfDatetime 日期时间字段名称序列,多个已英文逗号分开， 形如    deptname,username 
	 * @param isomer
	 * @return
	 */
	public static com.alibaba.fastjson.JSONObject isomerStyle(String text, Map<String, Object> varValues, String placeholder, String dateFormat, String datetimeFormat, String fieldsOfDatetime, com.alibaba.fastjson.JSONArray isomer) {
		com.alibaba.fastjson.JSONObject result = null;
		
		if (varValues != null && !varValues.isEmpty() && isomer != null && !isomer.isEmpty())
		{
			for (Entry<String, Object> entry : varValues.entrySet()) 
			{
				String key = entry.getKey();	
				Object value = entry.getValue();
				if (value != null)
				{
					for (int i=0; i<isomer.size(); i++)
					{					
						com.alibaba.fastjson.JSONObject iObject = isomer.getJSONObject(i);
						if (iObject != null)
						{
							String field = iObject.getString("field");
							if (field != null && field.trim().equalsIgnoreCase(key.trim()))
							{
								boolean isQueer = getIsomer(text, placeholder, field);
								if (isQueer)
								{//需要处理异构样式
									String it = iObject.getString("it");
									if (it != null)
									{
										com.alibaba.fastjson.JSONArray tar = com.alibaba.fastjson.JSONArray.parseArray(it);
										if (tar != null)
										{
											for (int i1=0; i1<tar.size(); i1++)
											{	
												com.alibaba.fastjson.JSONObject iTar = tar.getJSONObject(i1);
												if (iTar != null)
												{
													String enumeration = iTar.getString("enumeration");
													String logicalexpr = iTar.getString("logicalexpr");
													String style = iTar.getString("style");
													if (enumeration != null && !enumeration.trim().equals(""))
													{
														boolean conform = enumerationEval(enumeration, value);
														if (conform)
														{
															result = com.alibaba.fastjson.JSONObject.parseObject(style);
														}
														
													}
													else if (logicalexpr != null && !logicalexpr.trim().equals(""))
													{	
														Map expr = textAnaly(varValues, 1L, placeholder, dateFormat, datetimeFormat, fieldsOfDatetime, logicalexpr);
														if (expr != null)
														{
															logicalexpr = (String) expr.get("text");
															logicalexpr = logicalexprValid(logicalexpr, value);
															
															boolean conform = logicalexprEval(logicalexpr);
															if (conform)
															{
																result = com.alibaba.fastjson.JSONObject.parseObject(style);
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
				}							
			}			
		}
		
		return result;
	}
	
	/**
	 * 文本解析
	 * @param varValues 存放变量占位符及对应值的map
	 * @param isSimple 工作模式 0-正常模式（即用于msword中） 1-简单模式 （即直接用于字符串中）
	 * @param placeholder 数据占位符
	 * @param dateFormat 日期格式字符串 需要符合java.text.SampleDateFormat规范
	 * @param datetimeFormat 日期时间格式字符串 需要符合java.text.SimpleDateFormat规范
	 * @param fieldsOfDatetime 日期时间字段名称序列,多个已英文逗号分开， 形如    deptname,username 
	 * @param text 文本
	 * @return
	 */
	public static Map textAnaly(Map<String, Object> varValues, Long isSimple, String placeholder, String dateFormat, String datetimeFormat, String fieldsOfDatetime, String text) {
		Map result = new HashMap();
		
		boolean isSetText = false;
		if (text != null && !text.trim().equals("")) 
		{
			if (varValues != null)
			{
				for (Entry<String, Object> entry : varValues.entrySet()) 
				{
					String key = entry.getKey();								
					String regEx = "\\$\\{[\\s　  ]*" + key + "[\\s　  ]*\\}";
					if (placeholder != null && !placeholder.trim().equals(""))
					{
						regEx = "\\$\\{[\\s　  ]*" + placeholder.trim() + "\\.[\\s　  ]*" + key + "[\\s　  ]*\\}";
					}
					
					Pattern pattern = Pattern.compile(regEx, Pattern.CASE_INSENSITIVE);
					Matcher m = pattern.matcher(text);
					if (m.find()) {
						isSetText = true;
						Object value = entry.getValue();
						if (value != null)
						{
							if (value instanceof java.util.Date) 
							{												
								// 日期格式化替换
								if (fieldsOfDatetime != null && !fieldsOfDatetime.trim().equals(""))
								{
									String[] fieldsOf = fieldsOfDatetime.trim().split(",");
									if (fieldsOf != null && fieldsOf.length > 0) { } else
									{
										fieldsOf = fieldsOfDatetime.trim().split("，");
									}
									if (fieldsOf != null && fieldsOf.length > 0) { } else
									{
										fieldsOf = fieldsOfDatetime.trim().split(";");
									}
									if (fieldsOf != null && fieldsOf.length > 0) { } else
									{
										fieldsOf = fieldsOfDatetime.trim().split("；");
									}
									
									boolean isDatetime = false;
									for (int i4=0; i4<fieldsOf.length; i4++)
									{
										String fieldOf = fieldsOf[i4];
										if (fieldOf != null && !fieldOf.trim().equals(""))
										{
											if (fieldOf.trim().equalsIgnoreCase(key.trim()))
											{
												isDatetime = true;
											}															
										}														
									}																										
									if (isDatetime)
									{//作时间格式
										if (datetimeFormat != null && !datetimeFormat.trim().equals(""))
										{
										}
										else
										{
											datetimeFormat = "yyyy年MM月dd日 HH时mm分ss秒";
										}
										text = text.replaceAll(regEx, (new java.text.SimpleDateFormat(datetimeFormat.trim())).format((java.util.Date)value));
									}
									else
									{//作日期格式
										if (dateFormat != null && !dateFormat.trim().equals(""))
										{
										}
										else
										{
											dateFormat = "yyyy年MM月dd日";
										}
										text = text.replaceAll(regEx, (new java.text.SimpleDateFormat(dateFormat.trim())).format((java.util.Date)value));
										
									}
								}
								else
								{//作日期格式
									if (dateFormat != null && !dateFormat.trim().equals(""))
									{
									}
									else
									{
										dateFormat = "yyyy年MM月dd日";
									}
									text = text.replaceAll(regEx, (new java.text.SimpleDateFormat(dateFormat.trim())).format((java.util.Date)value));
									
								}
								
								System.out.println(text);
							
							}
							else// if (value instanceof String) 
							{										
								// 文本替换
								text = text.replaceAll(regEx, value.toString());
								
								System.out.println(text);
							}
						}
						else
						{
							//空字符处理	
							text = text.replaceAll(regEx, "");
							if (isSimple != null && isSimple.equals(1L))
							{
								text = "";
							}
							
						}
						
					}
					
					
				
				}
			}
		
		}
		
		result.put("isSetText", isSetText);
		result.put("text", text);

		return result;
	}
	
	/**
	 * 判断文本中是否包含字段名称
	 * @param text
	 * @param placeholder  数据占位符
	 * @param fieldName
	 * @return
	 */
	public static boolean getIsomer(String text, String placeholder, String fieldName){
		boolean result = false;
		if (text != null && !text.trim().equals("") && fieldName != null && !fieldName.trim().equals(""))
		{							
			String regEx = "\\$\\{[\\s　 ]*" + fieldName.trim() + "[\\s　 ]*\\}";
			if (placeholder != null && !placeholder.trim().equals(""))
			{
				regEx = "\\$\\{[\\s　 ]*" + placeholder + "\\." + fieldName.trim() + "[\\s　 ]*\\}";
			}
			
			Pattern pattern = Pattern.compile(regEx, Pattern.CASE_INSENSITIVE);
			Matcher m = pattern.matcher(text);
			if (m.find()) {
				result = true;				
			}			
			
		}		
		
		return result;
	}
	
	/**
	 * 判断变量值是否在枚举中
	 * @param enumeration 枚举
	 * @param value 变量值
	 * @return
	 */
	public static boolean enumerationEval(String enumeration, Object value) {
		boolean result = false;
		if (enumeration != null && !enumeration.trim().equals("") && value != null)
		{
			String[] tar = enumeration.trim().split(",");
			if (tar != null && tar.length > 0)
			{
				String valueStr = null;
				if (value instanceof Integer 
						|| value instanceof Long 
						|| value instanceof Double 
						|| value instanceof Float 
						|| value instanceof java.math.BigDecimal) 
				{
					valueStr = String.valueOf(value);
				}
				else if (value instanceof String) 
				{
					valueStr = (String) value;			
				}
				if (valueStr != null)
				{
					for (int i=0; i<tar.length; i++)
					{
						String iEnum = tar[i];
						if (iEnum != null && iEnum.trim().equalsIgnoreCase(valueStr.trim()))
						{
							result = true;
							break;
						}
						
					}	
				}
							
			}
		}
		return result;
	}
	
	/**
	 * 逻辑表达式有效化
	 * @param logicalexpr  逻辑表达式
	 * @param value 变量值
	 * @return
	 */
	public static String logicalexprValid(String logicalexpr, Object value) {
		String result = logicalexpr;
		if (logicalexpr != null && !logicalexpr.trim().equals("") && value != null)
		{
			java.text.SimpleDateFormat dateFormat = new java.text.SimpleDateFormat("yyyy-MM-dd");
			if (value instanceof java.util.Date) 
			{
				result = logicalexpr.replaceAll("\\?", "'" + dateFormat.format(value) + "'");
			}
			else if (value instanceof Integer 
					|| value instanceof Long 
					|| value instanceof Double 
					|| value instanceof Float 
					|| value instanceof java.math.BigDecimal) 
			{
				result = logicalexpr.replaceAll("\\?", String.valueOf(value));
			}
			else if (value instanceof String) 
			{
				result = logicalexpr.replaceAll("\\?", (String) value);			
			}
			
			result = result.replaceAll("\\{[\\s　 ]*sysdate[\\s　 ]*\\}", "'" + dateFormat.format(new java.util.Date()) + "'");
			result = result.replaceAll(" eq ", " == ");
			result = result.replaceAll(" gt ", " > ");
			result = result.replaceAll(" gte ", " >= ");
			result = result.replaceAll(" lt ", " < ");
			result = result.replaceAll(" lte ", " <= ");
		}
		
		return result;
	}
	
	/**
	 * 执行逻辑表达式获取结果
	 * @param logicalexpr 逻辑表达式
	 * @return
	 */
	public static boolean logicalexprEval(String logicalexpr) {
		boolean result = false;
		if (logicalexpr != null && !logicalexpr.trim().equals(""))
		{
			ScriptEngineManager manager = new ScriptEngineManager();
			ScriptEngine engine = manager.getEngineByName("JavaScript");
			
			Object target;
			try {
				target = engine.eval(logicalexpr);
				
				if (target instanceof java.lang.Boolean) 
				{					
					result = (java.lang.Boolean) target;
				}
				
			} catch (ScriptException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			
		}
		
		return result;
	}
	
	
	
	/**
	 * 抓取报表图片标签
	 * @param text 文本
	 * @return
	 */
	public static List<Picture> grabPictures(String text) {
		List<Picture> result = null;
		
		if (text != null && !text.trim().equals(""))
		{
			//解析beginTagIndex和beginTagText
			String regEx = "\\<[\\s]*picture [^\\>]*\\>";
			Pattern pattern = Pattern.compile(regEx, Pattern.CASE_INSENSITIVE);
			Matcher m = pattern.matcher(text);
			if (m.find()) {
				
				Picture picture = new Picture();
				picture.text = m.group(0);
				
				//解析field
				regEx = "field[\\s]*=[\\s]*[\"“”][\\s]*[^\"“”]*[\"“”]";
				pattern = Pattern.compile(regEx, Pattern.CASE_INSENSITIVE);
				m = pattern.matcher(picture.text);
				if (m.find()) {
					String tmpStr = m.group(0);
					tmpStr = tmpStr.replaceFirst("field[\\s]*=[\\s]*[\"“”]", "");							
					picture.field = tmpStr.replaceFirst("[\"“”]$", "");

				} else {
					regEx = "field[\\s]*=[\\s]*[^\\s\"“”]*[\\s]*";
					pattern = Pattern.compile(regEx, Pattern.CASE_INSENSITIVE);
					m = pattern.matcher(picture.text);
					if (m.find()) {
						String tmpStr = m.group(0);
						tmpStr = tmpStr.replaceFirst("field[\\s]*=[\\s]*", "");									
						picture.field = tmpStr;

					} else {
					}

				}	
				
				//解析type
				regEx = "type[\\s]*=[\\s]*[\"“”][\\s]*[^\"“”]*[\"“”]";
				pattern = Pattern.compile(regEx, Pattern.CASE_INSENSITIVE);
				m = pattern.matcher(picture.text);
				if (m.find()) {
					String tmpStr = m.group(0);
					tmpStr = tmpStr.replaceFirst("type[\\s]*=[\\s]*[\"“”]", "");							
					picture.type = tmpStr.replaceFirst("[\"“”]$", "");

				} else {
					regEx = "type[\\s]*=[\\s]*[^\\s\"“”]*[\\s]*";
					pattern = Pattern.compile(regEx, Pattern.CASE_INSENSITIVE);
					m = pattern.matcher(picture.text);
					if (m.find()) {
						String tmpStr = m.group(0);
						tmpStr = tmpStr.replaceFirst("type[\\s]*=[\\s]*", "");									
						picture.type = tmpStr;

					} else {
					}

				}	
				
				//解析width
				regEx = "width[\\s]*=[\\s]*[\"“”][\\s]*[^\"“”]*[\"“”]";
				pattern = Pattern.compile(regEx, Pattern.CASE_INSENSITIVE);
				m = pattern.matcher(picture.text);
				if (m.find()) {
					String tmpStr = m.group(0);
					tmpStr = tmpStr.replaceFirst("width[\\s]*=[\\s]*[\"“”]", "");	
					tmpStr = tmpStr.replaceFirst("[\"“”]$", "");
					int width = 100;
					if (tmpStr != null && !tmpStr.trim().equals(""))
					{
						try
						{
							width = Integer.valueOf(tmpStr);
						}
						catch(Exception e) {
							e.printStackTrace();
						}						
					}
					picture.width = width;

				} else {
					regEx = "width[\\s]*=[\\s]*[^\\s\"“”]*[\\s]*";
					pattern = Pattern.compile(regEx, Pattern.CASE_INSENSITIVE);
					m = pattern.matcher(picture.text);
					if (m.find()) {
						String tmpStr = m.group(0);
						tmpStr = tmpStr.replaceFirst("width[\\s]*=[\\s]*", "");									
						int width = 100;
						if (tmpStr != null && !tmpStr.trim().equals(""))
						{
							try
							{
								width = Integer.valueOf(tmpStr);
							}
							catch(Exception e) {
								e.printStackTrace();
							}						
						}
						picture.width = width;

					} else {
					}

				}	
				
				//解析height
				regEx = "height[\\s]*=[\\s]*[\"“”][\\s]*[^\"“”]*[\"“”]";
				pattern = Pattern.compile(regEx, Pattern.CASE_INSENSITIVE);
				m = pattern.matcher(picture.text);
				if (m.find()) {
					String tmpStr = m.group(0);
					tmpStr = tmpStr.replaceFirst("height[\\s]*=[\\s]*[\"“”]", "");	
					tmpStr = tmpStr.replaceFirst("[\"“”]$", "");
					int height = 100;
					if (tmpStr != null && !tmpStr.trim().equals(""))
					{
						try
						{
							height = Integer.valueOf(tmpStr);
						}
						catch(Exception e) {
							e.printStackTrace();
						}						
					}
					picture.height = height;

				} else {
					regEx = "height[\\s]*=[\\s]*[^\\s\"“”]*[\\s]*";
					pattern = Pattern.compile(regEx, Pattern.CASE_INSENSITIVE);
					m = pattern.matcher(picture.text);
					if (m.find()) {
						String tmpStr = m.group(0);
						tmpStr = tmpStr.replaceFirst("height[\\s]*=[\\s]*", "");									
						int height = 100;
						if (tmpStr != null && !tmpStr.trim().equals(""))
						{
							try
							{
								height = Integer.valueOf(tmpStr);
							}
							catch(Exception e) {
								e.printStackTrace();
							}						
						}
						picture.height = height;

					} else {
					}

				}	
				
				if (result == null)
				{
					result = new ArrayList<Picture>();
				}
				
				result.add(picture);

			} else {
				
			}		
		}
		
		return result;
	}
	
	/**
	 * 图片标签解析
	 * @param varValues 存放变量占位符及对应值的map
	 * @param placeholder 数据占位符
	 * @param runs 输入文本的在msword中的run列表
	 * @return
	 */
	public static Map pictureAnaly(Map<String, Object> varValues, String placeholder, List<XWPFRun> runs) {
		Map result = new HashMap();
		
		String text = "";
		if (runs != null && runs.size() > 0)
		{
			for (int j1=0; j1<runs.size(); j1++)
			{
				XWPFRun run = runs.get(j1);
				text += run.getText(0);
			}
			
			if (text != null && !text.trim().equals("")) 
			{
				List<Picture> pictures = grabPictures(text);
				if (pictures != null && pictures.size() > 0)
				{
					for (int i=0; i<pictures.size(); i++)
					{
						Picture picture = pictures.get(i);
						if (picture != null && picture.field != null && !picture.field.trim().equals(""))
						{
							int beginIndex = text.indexOf(picture.text);
							int endIndex = beginIndex + picture.text.length();
							
							System.out.println(text.substring(beginIndex, endIndex));
							
							String tmp_text = text.substring(0, beginIndex);
							int beginIndexOfRun = indexOfRunList(beginIndex, tmp_text, runs);							
							tmp_text += picture.text;
							int endIndexOfRun = indexOfRunList(endIndex - 1, tmp_text, runs);							
							
							String pictureFileName = null;						
							if (varValues != null && varValues.containsKey(picture.field.trim()))
							{
								Object value = varValues.get(picture.field.trim());
								if (value != null)
								{
									if (value instanceof String) 
									{
										pictureFileName = (String) value;
									}
								}
								boolean isExist = false;
								if (pictureFileName != null && !pictureFileName.trim().equals(""))
								{
									File picfile = new File(pictureFileName);
									if (picfile.exists())
									{
										isExist = true;
									}
								}
								
								if (isExist) { } else
								{
									pictureFileName = msword.Operate.class.getResource("pict_not_exist.jpg").getPath();																		
								}
								for (int ir=beginIndexOfRun; ir<=endIndexOfRun; ir++)
								{
									XWPFRun run = runs.get(ir);
									if (ir == beginIndexOfRun)
									{
										try {
											int type = Document.PICTURE_TYPE_JPEG;
											if (picture.type != null && picture.type.trim().equals("png"))
											{
												type = Document.PICTURE_TYPE_PNG;
											}
													
											run.addPicture(new FileInputStream(pictureFileName), type, pictureFileName, Units.toEMU(picture.width), Units.toEMU(picture.height));
										
										
										} catch (Exception e) {
											// TODO Auto-generated catch block
											e.printStackTrace();
										}
									}
									
									run.setText("", 0);
								}
								
								
								
								
								
							}
							
							
						}
						
						
					}
					
					result.put("pictures", pictures);
				}
				
			
			}
		}		
		
		
		

		return result;
	}


	
}
