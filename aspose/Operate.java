package aspose;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import msword.Report;

public class Operate {
	
	private static final Logger LOG = LoggerFactory.getLogger(Operate.class);
	/**
	 * 另保存pdf工作文件
	 * @param report
	 * @return
	 */
	
	public static boolean convertasPdf(Report report) {
		boolean result = false;
		if (report != null 
				&& report.workFilename != null
				&& !report.workFilename.trim().equals(""))
		{
			String licenseName = Thread.currentThread().getContextClassLoader().getResource("").getPath() + "aspose-words-license.xml";
			
			com.aspose.words.License license = new com.aspose.words.License();
			try {
				license.setLicense(licenseName);
			} catch (Exception e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}
			
			
			
			
			File docxFile = new File(report.workFilename.trim());
			String pdfFilename = report.workFilename.trim() + ".pdf";			
			int dpos = report.workFilename.trim().lastIndexOf(".docx");
			if (dpos > -1)
			{
				pdfFilename = pdfFilename.substring(0, dpos) + ".pdf";
			}			
			File pdfFile = new File(pdfFilename);
			
			InputStream docxStream = null;
			OutputStream pdfStream = null;
			try {
				if (docxFile.exists())
				{
					docxStream = new FileInputStream(docxFile);
					pdfStream = new FileOutputStream(pdfFile);
					
					com.aspose.words.Document doc = new com.aspose.words.Document(docxStream);
					com.aspose.words.PdfSaveOptions pdfSaveOptions = new com.aspose.words.PdfSaveOptions();
					pdfSaveOptions.setSaveFormat(com.aspose.words.SaveFormat.PDF);
					
					doc.save(pdfStream, pdfSaveOptions);
					docxStream.close();
					pdfStream.flush();
					pdfStream.close();
					
					report.pdfFilename = pdfFilename;
					
					result = true;
				}
			} catch (Exception e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			finally {				
				docxFile = null;
				pdfFile = null;
				if (docxStream != null)
				{
					try {
						docxStream.close();
					} catch (IOException e) {
						LOG.error("",e);
					}
					docxStream = null;
				}
				if (pdfStream != null)
				{
					try {
						pdfStream.close();
					} catch (IOException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
					pdfStream = null;
				}
				
			}
		}
		
		return result;
	}
	
	/**
	 * 另保存html工作文件
	 * @param report
	 * @return
	 */
	public static boolean convertasHtml(Report report) {
		boolean result = false;
		if (report != null 
				&& report.workFilename != null
				&& !report.workFilename.trim().equals(""))
		{
			String licenseName = Thread.currentThread().getContextClassLoader().getResource("").getPath() + "aspose-words-license.xml";
			
			com.aspose.words.License license = new com.aspose.words.License();
			try {
				license.setLicense(licenseName);
			} catch (Exception e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}
			
			
			
			
			File docxFile = new File(report.workFilename.trim());
			String htmlName = report.workFilename.trim() + ".html";			
			int dpos = report.workFilename.trim().lastIndexOf(".docx");
			if (dpos > -1)
			{
				htmlName = htmlName.substring(0, dpos) + ".html";
			}
			InputStream docxStream = null;
			try {
				if (docxFile.exists())
				{
					docxStream = new FileInputStream(docxFile);

					com.aspose.words.Document doc = new com.aspose.words.Document(docxStream);
					com.aspose.words.HtmlSaveOptions htmlSaveOptions = new com.aspose.words.HtmlSaveOptions();
					
					htmlSaveOptions.setSaveFormat(com.aspose.words.SaveFormat.HTML);
					
					doc.save(htmlName, htmlSaveOptions);
					docxStream.close();
					
					result = true;
				}
			} catch (Exception e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			finally {				
				docxFile = null;
				if (docxStream != null)
				{
					try {
						docxStream.close();
					} catch (IOException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
					docxStream = null;
				}
				
			}
		}
		
		return result;
	}
	
	/**
	 * 另保存xps工作文件
	 * @param report
	 * @return
	 */
	public static boolean convertasXps(Report report) {
		boolean result = false;
		if (report != null 
				&& report.workFilename != null
				&& !report.workFilename.trim().equals(""))
		{
			String licenseName = Thread.currentThread().getContextClassLoader().getResource("").getPath() + "aspose-words-license.xml";
			
			com.aspose.words.License license = new com.aspose.words.License();
			try {
				license.setLicense(licenseName);
			} catch (Exception e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}
			
			
			
			
			File docxFile = new File(report.workFilename.trim());
			String xpsName = report.workFilename.trim() + ".xps";			
			int dpos = report.workFilename.trim().lastIndexOf(".docx");
			if (dpos > -1)
			{
				xpsName = xpsName.substring(0, dpos) + ".xps";
			}
			File xpsFile = new File(xpsName);
			
			InputStream docxStream = null;
			OutputStream xpsStream = null;
			try {
				if (docxFile.exists())
				{
					docxStream = new FileInputStream(docxFile);
					xpsStream = new FileOutputStream(xpsFile);
					
					com.aspose.words.Document doc = new com.aspose.words.Document(docxStream);
					com.aspose.words.XpsSaveOptions xpsSaveOptions = new com.aspose.words.XpsSaveOptions();
					
					xpsSaveOptions.setSaveFormat(com.aspose.words.SaveFormat.XPS);
					
					doc.save(xpsStream, xpsSaveOptions);
					docxStream.close();
					xpsStream.flush();
					xpsStream.close();
					
					result = true;
				}
			} catch (Exception e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			finally {				
				docxFile = null;
				xpsFile = null;
				if (docxStream != null)
				{
					try {
						docxStream.close();
					} catch (IOException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
					docxStream = null;
				}
				if (xpsStream != null)
				{
					try {
						xpsStream.close();
					} catch (IOException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
					xpsStream = null;
				}
				
			}
		}
		
		return result;
	}
	
	/**
	 * 另保存png工作文件
	 * @param report
	 * @return
	 */
	public static boolean convertasPng(Report report) {
		boolean result = false;
		if (report != null 
				&& report.workFilename != null
				&& !report.workFilename.trim().equals(""))
		{
			String licenseName = Thread.currentThread().getContextClassLoader().getResource("").getPath() + "aspose-words-license.xml";
			
			com.aspose.words.License license = new com.aspose.words.License();
			try {
				license.setLicense(licenseName);
			} catch (Exception e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}
			
			
			
			
			File docxFile = new File(report.workFilename.trim());
			String pngName = report.workFilename.trim();
			int dpos = report.workFilename.trim().lastIndexOf(".docx");
			if (dpos > -1)
			{
				pngName = pngName.substring(0, dpos);
			}
			
			InputStream docxStream = null;
			try {
				if (docxFile.exists())
				{
					docxStream = new FileInputStream(docxFile);
					
					com.aspose.words.Document doc = new com.aspose.words.Document(docxStream);
					com.aspose.words.ImageSaveOptions imageSaveOptions= new com.aspose.words.ImageSaveOptions(com.aspose.words.SaveFormat.PNG);
					imageSaveOptions.setPrettyFormat(true);
					imageSaveOptions.setUseAntiAliasing(true);
					imageSaveOptions.setJpegQuality(100);
					for (int i=0;  i<doc.getPageCount();  i++)
					{
						imageSaveOptions.setPageIndex(i);
						doc.save(pngName + "-" + i + ".png", imageSaveOptions);
					}					
					docxStream.close();
					
					result = true;
				}
			} catch (Exception e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			finally {				
				docxFile = null;
				if (docxStream != null)
				{
					try {
						docxStream.close();
					} catch (IOException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
					docxStream = null;
				}

			}
		}
		
		return result;
	}
	
}
