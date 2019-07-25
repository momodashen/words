package msword;

import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;
import java.util.Map;

import org.apache.poi.xwpf.usermodel.XWPFDocument;

import com.huazi.common.dao.SmartBaseDao;

import msword.report.Root;
import msword.report.Table;

public class Report {
	
	public String tmpFilename; //msword文件物理地址作为模板文件
	public InputStream is; //输入文件流
	public XWPFDocument document; //msword文档
	public String workFilename; //msword文件物理地址作为工作文件
	public String pdfFilename; //pdf文件物理地址作为工作文件
	public OutputStream os; //输出文件流
	
	public List<Root> rootList; //root数据的定义列表
	
	public List<Table> tables; //报表msword表格列表
	public Map<String, Object> varValues; //报表数据map容器
	public 	SmartBaseDao smartBaseDao; 

		
	/**
	 * 解析
	 */
	public void transformDOCX() {
		
		//载入模板文件
		msword.report.Operate.load(this);
		//解析
		msword.report.Operate.execute(this);		
		//再替换没有匹配的${}为空字符
		msword.report.Operate.scavengeSuperfluous(this);
		//保存工作文件
		msword.report.Operate.save(this);
		//转换为pdf文件
		aspose.Operate.convertasPdf(this);
		//转换为html文件
//		aspose.Operate.convertasHtml(this);
		//转换为xps文件
//		aspose.Operate.convertasXps(this);
		//转换为png文件
//		aspose.Operate.convertasPng(this);
	}

}
