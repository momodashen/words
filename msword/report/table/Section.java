package msword.report.table;

import java.util.ArrayList;
import java.util.List;


import msword.report.Table;
import msword.report.table.section.SectionRow;
import msword.report.table.section.TemplateRow;

/**
 * 报表msword表格中的一块数据区域
 *
 */
public class Section {
	
	public Table table; //所属的msword表 
	public int lever = -1; //层深
	public int selfIndex = -1; //在列表中的索引
	public List<Section> children; //子节点

	public int beginIndex = -1; //开始索引
	public int endIndex = -1; //结束索引
	public String attributeText = null; //属性文本
	
	public String alt; //当数据源无数据时,用该文本替换
	public int colTitleIndex = -1; //列标题行索引
	public String sql; //取数据的sql语句	
	public String dataSourceKey = "items"; //数据源的名称
	public String rowPlaceholder = "item"; //数据行占位符
	public int colOfSerialNumber = -1; //序号所在的列索引 索引以0开始 第一列为0
	public String dateFormat = "yyyy-MM-dd"; //日期格式字符串 需要符合java.text.SampleDateFormat规范
	public String datetimeFormat = "yyyy-MM-dd HH:mm:ss"; //日期时间格式字符串 需要符合java.text.SimpleDateFormat规范
	public String fieldsOfDatetime = ""; //日期时间字段名称序列,多个已英文逗号分开， 形如    deptname,username 
	public String fieldsOfMergerow = ""; //需要合并行数据的列字段名称序列,多个已英文逗号分开， 形如    deptname,username 
	
	public com.alibaba.fastjson.JSONArray isomer = null; //异构样式配置 json对象 字符串格式，形如isomer="[{field:'x13',it:[{logicalexpr:'(? + ${item.x14}) eq 2',style:{color:'97FFFF'}}]},{field:'x15',it:[{logicalexpr:'? eq {sysdate}',style:{color:'97FFFF'}}]}]"
																																				//	   ，形如isomer="[{field:'x13',it:[{enumeration:'a,b,c',style:{color:'97FFFF'}}]},{field:'x15',it:[{enumeration:'e,f,g',style:{color:'97FFFF'}}]}]"

	public List<TemplateRow> rptRowTmp = new ArrayList<TemplateRow>(); //行模板		
	public List<SectionRow> appendRptRows = new ArrayList<SectionRow>(); //根据数据源的记录数添加的表格行索引列表


}
