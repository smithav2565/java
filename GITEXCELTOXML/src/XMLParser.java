

import java.io.File;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.DocumentBuilder;
import org.w3c.dom.Document;
import org.w3c.dom.NodeList;
import org.w3c.dom.Node;
import org.w3c.dom.Element;


import org.apache.log4j.Logger;

public class XMLParser {
	
	public static Logger logger = Logger.getLogger("XMLParser");

	public static void parseXML(File xmlFile){
		
		try{
		    SAXParserFactory saxParserFactory = SAXParserFactory.newInstance();
		        SAXParser saxParser = saxParserFactory.newSAXParser();
		        saxParser.parse(xmlFile, new DataHandler());
		}
		catch(Exception exp){
			logger.debug("***********************88"+exp.getMessage());
			exp.printStackTrace();

	}
		
	}

}
