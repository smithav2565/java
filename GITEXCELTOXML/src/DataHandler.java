

import java.util.ArrayList;
import java.util.Arrays;

import org.apache.log4j.Logger;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

public class DataHandler extends DefaultHandler {

	public static Logger logger = Logger.getLogger("LoanHandler");

	private boolean isNewRec = false;
	private boolean isExcludeTag = false;
	private ArrayList<String> excludeTagList = new ArrayList<String>(Arrays.asList(ConverterMain.prop.getProperty("EXCLUDE_TAGS").split(",")));

	private String key = "";
	private String value = "";
	private String rowCreatorTag = ConverterMain.prop.getProperty("ROW_CREATOR_TAG");
	private static int recCnt = -1;

	
	  
	@Override
	    public void startElement(String uri, String localName, String qName, Attributes attributes)
	            throws SAXException {
    	if(qName.equalsIgnoreCase(rowCreatorTag.trim()) ){
    		logger.debug("loan id count: "+recCnt);
    		recCnt++;

    	}
    	
	    	
	    	
	    	key = qName;
	     
	    }

	    @Override
	    public void endElement(String uri, String localName, String qName) throws SAXException {
	    	
	    	if(!excludeTagList.contains(qName)){

		    ExcelCreator.writeToExcel(qName, value, recCnt);
	    	}
		    	
	    	value = "";
	    	
	    
	    }
	    
	  

	    @Override
	    public void characters(char ch[], int start, int length) throws SAXException {
	    	
	    	   		value = value.concat(new String(ch, start, length));
	    	
	    	
	    	

	  
	    }
	    


}
