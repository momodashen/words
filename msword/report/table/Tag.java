package msword.report.table;

import msword.report.Table;
import msword.report.table.tag.Type;

/**
 * 报表表格数据行开始及结束标签
 *
 */
public class Tag {
	
	public Table table; //所属的msword表 
	
	public Type type = Type.NoneOf; //类别 开始标签或结束标签  0 开始标签 1 结束标签	
	public int index = -1; //标签在msword表格中的行索引
	public String text; //标签字符串文本
	
	public int lever = -1; //层深
	public int pairingIndex = -1; //对偶项在msword表格中的行索引
	public boolean isOutlier = false; //是否孤立点

}
