package msword.report.table.section;

import java.util.List;

import org.apache.poi.xwpf.usermodel.XWPFTableRow;

/**
 * 数据区域行模板单元
 *
 */
public class TemplateRow {
	
	public XWPFTableRow tmpRow; //行模板单元关联的msword中的行
	public List<Column> columnList; //模板单元列列表
	public int index; //行模板单元关联的msword中的行索引

}
