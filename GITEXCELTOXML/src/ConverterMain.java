

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;
import java.util.Properties;

import org.apache.log4j.Logger;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



public class ConverterMain {
	public static Properties prop = null;
	public static String OLD_XLS_TO_MERGE = "";
	public static Logger logger = Logger.getLogger("ConverterMain");

	public static void main(String[] args) throws OpenXML4JException {
		if(args.length == 0){
			logger.debug("Input XML folder to be parsed");
	        System.exit(0);

		}
		loadProp();


		//String strInputXmlPath = "G:\\MBA docs\\Projects\\Kiva\\kiva_ds_xml\\loans";
		String strInputXmlPath = args[0];
		File folder = new File(strInputXmlPath);
		File[] xmlFiles = folder.listFiles();
		int fileCnt = 0;
		for(File xmlFile: xmlFiles){
			logger.debug("Parsing xml: "+xmlFile.getName());
		//XMLParser.parseXML(new File("data.xml"));
		XMLParser.parseXML(xmlFile);	
		logger.debug("Parsing xml: "+xmlFile.getName() +" done.");
			fileCnt++;

			

		}
		logger.debug("************Total number of files parsed**************: "+fileCnt);

		
		logger.debug("writing workbook to excel..");
		ExcelCreator.writeWorkBookTofile();
		logger.debug("writing workbook to excel done.");
		
		XLSXCopy xlsxcopy = new XLSXCopy();
		try {
			xlsxcopy.processAllSheets(ConverterMain.prop.getProperty("XSL_FILE_NAME").concat(ExcelCreator.TIME_STAMP).concat(".xlsx"));
			File tempFile = new File(ConverterMain.prop.getProperty("XSL_FILE_NAME").concat(ExcelCreator.TIME_STAMP).concat(".xlsx"));
			tempFile.delete();

		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		
	
	}
	
	/**
	 * 
	 */
	/**
	 * 
	 */
	private static void loadProp(){
		prop = new Properties();
		InputStream input = null;

		try {

			input = new FileInputStream("src\\config.properties");

			// load a properties file
			prop.load(input);

			

		} catch (IOException ex) {
			ex.printStackTrace();
		} finally {
			if (input != null) {
				try {
					input.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}

		
	}
	}

}
