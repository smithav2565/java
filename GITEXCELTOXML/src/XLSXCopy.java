
import java.io.FileOutputStream;
import java.io.InputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.TimeZone;

import org.apache.log4j.Logger;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.xml.sax.Attributes;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

public class XLSXCopy {
	public static Logger logger = Logger.getLogger("XLSXCopy");
	private static SXSSFWorkbook  workbook = new SXSSFWorkbook();
	private static SXSSFSheet  currentSheet;
	private ArrayList<String> headerList;


	public void processAllSheets(String filename) throws Exception {
		
		OPCPackage pkg = OPCPackage.open(filename, PackageAccess.READ);
		try {
			XSSFReader r = new XSSFReader(pkg);
			SharedStringsTable sst = r.getSharedStringsTable();

			XMLReader parser = fetchWorkbookSheetParser(sst);

			Iterator<InputStream> sheets = r.getSheetsData();
			while (sheets.hasNext()) {
				currentSheet = (SXSSFSheet)workbook.createSheet();
				Row row = currentSheet.createRow(0);
				headerList = ExcelCreator.getHeaderList();
				int cnt =0;
				for(String key: headerList){
					row.createCell(cnt++).setCellValue(key);
					
				}
				
				InputStream sheet = sheets.next();
				InputSource sheetSource = new InputSource(sheet);
				parser.parse(sheetSource);
				sheet.close();
			}
			try {
				FileOutputStream  out = new FileOutputStream(ConverterMain.prop.getProperty("XLSX_NAME"));
	            workbook.write(out);
	            out.flush();
	            out.close();
	           // TIME_STAMP = "";
	        } catch (Exception e) {
	        	logger.debug(e.getMessage());
	        }
		} finally {
			pkg.close();
		}
	}

	public XMLReader fetchWorkbookSheetParser(SharedStringsTable sst) throws SAXException {
		XMLReader parser = XMLReaderFactory.createXMLReader();
		ContentHandler handler = new workbookSheetHandler(sst);
		parser.setContentHandler(handler);
		return parser;
	}

	private static class workbookSheetHandler extends DefaultHandler {
		private final SharedStringsTable sst;
		private String contents;
		private boolean nextIsString;
		private boolean inlineStr;
		private  int RowCnt = 0;
		private  int ColCnt = 0;

	
		
		private workbookSheetHandler(SharedStringsTable sst) {
			this.sst = sst;
		}

		@Override
        public void startElement(String uri, String localName, String name,
								 Attributes attributes) throws SAXException {
			// c => cell
			if(name.equals("row")){
				RowCnt++;
				XLSXCopy.currentSheet.createRow(RowCnt);
				ColCnt = 0;
			}
			if(name.equals("c")) {
				
				String cellType = attributes.getValue("t");
				nextIsString = cellType != null && cellType.equals("s");
				inlineStr = cellType != null && cellType.equals("inlineStr");
			}
			contents = "";
		}

		@Override
        public void endElement(String uri, String localName, String name)
				throws SAXException {

			if(nextIsString) {
				Integer idx = Integer.valueOf(contents);
				nextIsString = false;
			}

			if(name.equals("v") || (inlineStr && name.equals("c"))) {
				setCellValByType(XLSXCopy.currentSheet.getRow(RowCnt).createCell(ColCnt++), contents.trim());
				//XLSXCopy.currentSheet.getRow(RowCnt).createCell(ColCnt++).setCellValue(lastContents);
			}
		}
		private  void setCellValByType(Cell cell,String value){
			   if(isValidDate(value)){
				   cell.setCellValue(value);
			   }
			   else if(value.matches("^-?\\d+(\\.\\d+)?$") && value.trim().length()>0){
				   
				  // if(value.indexOf(".") > -1){

				   cell.setCellValue(Double.parseDouble(value));
				 //  }
				 ////  else{
					 //  cell.setCellValue(Integer.parseInt(value));

				  // }
			   }
			   else{
				   cell.setCellValue(value);
			   }
			   
		   }
		   
		   private  boolean isValidDate(String strDate){
			   boolean isValidDate = true;
			   SimpleDateFormat sdf = new SimpleDateFormat(ConverterMain.prop.getProperty("DATE_FORMAT"));
			   sdf.setTimeZone(TimeZone.getTimeZone(ConverterMain.prop.getProperty("TIME_ZONE")));
			   try{
				   sdf.parse(strDate);
			   }
			   catch(ParseException exp){
				   isValidDate = false;
			   }
			   return isValidDate;

			   
			   
		   }
		    


		@Override
        public void characters(char[] ch, int start, int length) throws SAXException { 
			contents += new String(ch, start, length);
		}
	}


}