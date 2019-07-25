package msword.report.table.section;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.xwpf.usermodel.XWPFTable;

/**
 * 列
 *
 */
public class Column {
	
	public String text = ""; //文本内容
	public String field = ""; //数据字段名
	public List<XWPFTable> xwpfTableList = new ArrayList<XWPFTable>(); //内含msword表格列表
	public int index = -1; //索引
	
	
	 
}
