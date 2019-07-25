package msword.report;

import java.util.ArrayList;
import java.util.List;
import org.apache.poi.xwpf.usermodel.XWPFTable;

import msword.Report;
import msword.report.table.Section;
import msword.report.table.Tag;

/**
 * 报表msword表格
 *
 */
public class Table {
	
	public Report report; //报表 对象
	public int xwpfTableIndex; //poi表格索引
	public XWPFTable xwpfTable; //poi表格
	public List<Integer> cacheIndexList = new ArrayList<Integer>(); //用于缓存数据解析时msword中表格的行row索引
	
	public List<Tag> tagList; //报表表格数据行开始及结束标签列表	
	public List<Section> sectionList; //数据区域列表	
	public List<Section> sectionTree; //数据区域树	
	public int levelCount = 0; //开始及标签tag的层深	
	
	

}
