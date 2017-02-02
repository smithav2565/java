

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.sql.Timestamp;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.TimeZone;

import org.apache.poi.xssf.eventusermodel.*;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;

import org.apache.log4j.Logger;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler;
import org.apache.poi.xssf.model.MapInfo;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFMap;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;






public class ExcelCreator {
	public static Logger logger = Logger.getLogger("ExcelCreator");

	public static String TIME_STAMP = "";
    private static final SimpleDateFormat sdf = new SimpleDateFormat("yyyy_MM_dd_HH_mm_ss");
    public static SXSSFWorkbook workbook;
    private static int prevRecCnt = 0;
   private static Sheet currentSheet;
	private static Row firstRow;
   private static int rowCnt = 0;
	private static ArrayList<String> headerList = new ArrayList<String>();
	
	public static ArrayList<String> getHeaderList(){
		return headerList;
	}
	
	public static void writeToExcel(String key, String value, int recCnt){
		loadExcelFile();
    

    	 
   SXSSFRow valueRow = null;
   if(null ==firstRow){
	   firstRow = currentSheet.createRow(0);
	  // currentSheetFirstRow = firstRow;
	   valueRow =(SXSSFRow) firstRow;
   }

   
  
   

   
   if(recCnt == prevRecCnt ){
		   valueRow = (SXSSFRow)currentSheet.getRow(rowCnt);	

   }
   else{

	   logger.debug("loan id: "+value);

	    valueRow =(SXSSFRow)createNewRow();
	    
	    prevRecCnt =0;
	   
   }
   Cell cell;
   boolean keyFound = false;
   int columnIndex = -1;
   for(String headerKey : headerList){
	   columnIndex++;
	   if(headerKey.equalsIgnoreCase(key.trim())){
		   keyFound = true;
		   break;
	   }
   }
    
   if(!keyFound){
	   headerList.add(key.trim());
	   columnIndex = headerList.size()-1 ;
   }


   setCellValByType(valueRow.createCell(columnIndex), value);
   
   prevRecCnt = recCnt;
  
    
    }
    
    
   private static void setCellValByType(Cell cell,String value){
	   if(isValidDate(value)){
		   cell.setCellValue(value);
	   }
	   else if(value.matches("^-?\\d+(\\.\\d+)?$") && value.trim().length()>0){		   

		   cell.setCellValue(Double.parseDouble(value));
	   }
	   else{
		   cell.setCellValue(value);
	   }
	   
   }
   
   private static boolean isValidDate(String strDate){
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
    
    private static Row createNewRow(){
    	Row row;
    	if(((currentSheet.getLastRowNum()))  >= SpreadsheetVersion.EXCEL2007.getMaxRows()-5){
			currentSheet = workbook.createSheet();
			
			workbook.setActiveSheet(workbook.getNumberOfSheets()-1);
			rowCnt =1;
			row =  currentSheet.createRow(1);
		}
    	else{
    		rowCnt++;
    		row =  currentSheet.createRow(rowCnt);
    	}
    
    	
    	return row;
    	
    }
    
    
	  	    
	    private static Sheet loadExcelFile(){
	    	if(TIME_STAMP.length() > 0 ){
	    		currentSheet = workbook.getSheetAt(workbook.getActiveSheetIndex());	    		    		
	    	}
	    	else{
	    		TIME_STAMP = sdf.format(new Timestamp(System.currentTimeMillis()));
	    		workbook = new SXSSFWorkbook(); 
	    		currentSheet = workbook.createSheet();
    			workbook.setActiveSheet(workbook.getNumberOfSheets()-1);
	    		
	    	}
	    	return currentSheet;
	    	
	    }
	    
	  public static void writeWorkBookTofile(){
		  FileOutputStream out;

	        try {
	             out = new FileOutputStream(ConverterMain.prop.getProperty("XSL_FILE_NAME").concat(TIME_STAMP).concat(".xlsx"));
	            workbook.write(out);
	            out.flush();
	            out.close();
	           // TIME_STAMP = "";
	        } catch (Exception e) {
	        	logger.debug(e.getMessage());
	        }
	       
	        

		  
	  }
	 

}
