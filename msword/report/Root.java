package msword.report;

import java.util.Map;

import org.apache.poi.xwpf.usermodel.XWPFTable;


/**
 * 报表root数据的定义对象
 *
 */
public class Root {
	
	public XWPFTable xwpfTable; //所属的msword表 
	
	public int selfIndex = -1; //在列表中的索引
	public int index = -1; //标签在msword表格中的行索引
	public String text; //标签字符串文本
	
	public String sql; //取数据的sql语句
	public String placeholder = "root"; //root数据占位符
	public String dateFormat = "yyyy年MM月dd日"; //日期格式字符串 需要符合java.text.SampleDateFormat规范
	public String datetimeFormat = "yyyy年MM月dd日 HH时mm分ss秒"; //日期时间格式字符串 需要符合java.text.SimpleDateFormat规范
	public String fieldsOfDatetime = ""; //日期时间字段名称序列,多个已英文逗号分开， 形如    deptname,username 
	
	public com.alibaba.fastjson.JSONArray isomer = null; //异构样式配置 json对象 字符串格式，形如isomer="[{field:'x13',it:[{logicalexpr:'(? + ${item.x14}) eq 2',style:{color:'97FFFF'}}]},{field:'x15',it:[{logicalexpr:'? eq {sysdate}',style:{color:'97FFFF'}}]}]"

	public Map<String, Object> varValues; //数据区域数据map容器

}
