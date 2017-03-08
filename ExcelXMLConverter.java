import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import java.io.File;
import java.io.FileOutputStream;
import java.net.URL;


/**
 * Convert XML file to Excel Spreadsheet
 * @imports Apache POI, Apache Commons, W3C libraries
 * @author Rebecca Ramnauth
 * Date: 3-3-2017
 */
public class ExcelXMLConverter {
    private static Workbook workbook;
    private static int rowNum;

    private final static int TITLE_NUM_COLUMN = 0;
    private final static int TITLE_SUBJECT = 1;
    private final static int TITLE_PARTS = 2;
    private final static int TITLE_REVISION_DATE = 3;
    private final static int TITLE_CONTAINS = 4;
    private final static int TITLE_DATE_PUBLISHED = 5;  
    private final static int SUBTITLE_NAME_COLUMN = 6;
    private final static int CHAPTER_NUM_COLUMN = 7;
    private final static int PART_NUM_COLUMN = 8;
    private final static int PART_NAME_COLUMN = 9;
    private final static int SUBPART_COLUMN = 10;
    private final static int SECTION_NUM_COLUMN = 11;
    private final static int SECTION_NAME_COLUMN = 12;
    private final static int SUBREQ_PARA_COLUMN = 13;
    private final static int CITA_COLUMN = 14;

    public static void main(String[] args) throws Exception {
        retrieveAndReadXml();
    }

    /**
     *
     * Downloads/Finds a XML file, reads the contents and then writes them to rows on an excel file.
     * @throws Exception
     */
    private static void retrieveAndReadXml() throws Exception {
        System.out.println("Completed read of XML file");

        /* File xmlFile = File.createTempFile("substances", "tmp");
         * String xmlFileUrl = "http://whatever";
         * URL url = new URL(xmlFileUrl);
         * System.out.println("Downloading file from " + xmlFileUrl + " ...");
         * FileUtils.copyURLToFile(url, xmlFile);
         * System.out.println("Downloading COMPLETED. Parsing...");

         If the XML file is stored externally, uncomment the above code. 
         */
        File xmlFile = new File("CFR-2016-title49-vol8.xml");

        initXls();

        Sheet sheet = workbook.getSheetAt(0);

        DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
        DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
        Document doc = dBuilder.parse(xmlFile);
		
		NodeList tList = doc.getElementsByTagName("TITLEPG");
		for (int a = 0; a < tList.getLength(); a++){
			System.out.println("Processing TITLE " + (a + 1) + "/" + tList.getLength());
			Node home = tList.item(a);
			if (home.getNodeType() == Node.ELEMENT_NODE){
				Element title = (Element) home;
				String title_num, subject, included_parts, revision_date, contain_descrip, date_published;
				
				try {
					title_num = title.getElementsByTagName("TITLENUM").item(0).getTextContent();
				} catch (NullPointerException e){
					title_num = "";
				}
				try {
					subject = title.getElementsByTagName("SUBJECT").item(0).getTextContent();
				} catch (NullPointerException e){
					subject = "";
				}
				try {
					included_parts = title.getElementsByTagName("PARTS").item(0).getTextContent();
				} catch (NullPointerException e){
					included_parts = "";
				}
				try {
					revision_date = title.getElementsByTagName("REVISED").item(0).getTextContent();
				} catch (NullPointerException e){
					revision_date = "";
				}
				try {
					contain_descrip = title.getElementsByTagName("CONTAINS").item(0).getTextContent();
				} catch (NullPointerException e){
					contain_descrip = "";
				}
				try {
					date_published = title.getElementsByTagName("DATE").item(0).getTextContent();
				} catch (NullPointerException e){
					date_published = "";
				}
				
				NodeList subList = doc.getElementsByTagName("SUBTITLE");
				for (int b = 0; b < subList.getLength(); b++){
					System.out.println("Processing SUBTITLE " + (b + 1) + "/" + subList.getLength());
					Node subhome = subList.item(b);
					if (subhome.getNodeType() == Node.ELEMENT_NODE){
						Element subtitle = (Element) subhome;
						String hd_source;
						
						try{
							hd_source = subtitle.getElementsByTagName("HD").item(0).getTextContent();
						} catch (NullPointerException e){
							hd_source = "";
						}
						
						//NodeList cList = subtitle.getElementsByTagName("CHAPTER");
						//for (int c = 0; c < cList.getLength(); c++){
						//System.out.println("Processing CHAPTER " + (c + 1) + "/" + cList.getLength());
						//Node chap = cList.item(c);
						//if (chap.getNodeType() == Node.ELEMENT_NODE){
							//Element chapter = (Element) chap;
						
					        NodeList nList = doc.getElementsByTagName("SUBPART");
					        for (int i = 0; i < nList.getLength(); i++) {
					            System.out.println("Processing SUBPART " + (i + 1) + "/" + nList.getLength());
					            Node node = nList.item(i);
					            if (node.getNodeType() == Node.ELEMENT_NODE) {
					                Element element = (Element) node;
					                String header, section_num, sb_subject;
					                
					                try {
					                	header = element.getElementsByTagName("HD").item(0).getTextContent();
					                } catch (NullPointerException e){
					                	header = " ";
					                }
					
					                NodeList prods = element.getElementsByTagName("SECTION");
					                for (int j = 0; j < prods.getLength(); j++) {
					                	System.out.println("	Processing SECTION " + (j + 1) + "/" + prods.getLength());
					                    Node prod = prods.item(j);
					                    if (prod.getNodeType() == Node.ELEMENT_NODE) {
					                        Element product = (Element) prod;
					                        String ss_num, ss_subject, citation;
					                        
					                        try{
					                        	ss_num = product.getElementsByTagName("SECTNO").item(0).getTextContent();
					                        } catch (NullPointerException e){
					                        	ss_num = "";
					                        }
					                        try{
					                        	ss_subject = product.getElementsByTagName("SUBJECT").item(0).getTextContent();
					                        } catch (NullPointerException e){
					                        	ss_subject = "";
					                        }
					                        try{
					                        	citation = product.getElementsByTagName("CITA").item(0).getTextContent();
					                        } catch (NullPointerException e){
					                        	citation = "";
					                        }
					                        
					                        NodeList sub_requirements = product.getElementsByTagName("P");
					                        for (int k = 0; k < sub_requirements.getLength(); k++){
					                        	System.out.println("		Processing SUB-REQUIREMENT " + (k + 1) + "/" + sub_requirements.getLength());
					                        	Node req = sub_requirements.item(k);
					                        	if (req.getNodeType() == Node.ELEMENT_NODE){
					                        		Element requirement = (Element) req;
					                        		String ss_para = requirement.getTextContent();
					                        		
					                        		Row row = sheet.createRow(rowNum++);
					                    
							                        //TITLE DETAILS
							                        Cell cell = row.createCell(TITLE_NUM_COLUMN);
							                        if (title_num != null && !title_num.equals(""))
							                        	cell.setCellValue(title_num);
							                        else
							                        	cell.setCellValue(" ");
							                        
							                        cell = row.createCell(TITLE_SUBJECT);
							                        if (subject != null && !subject.equals(""))
							                        	cell.setCellValue(subject);
							                        else
							                        	cell.setCellValue(" ");
							                        
							                        cell = row.createCell(TITLE_PARTS);
							                        if (included_parts != null && !included_parts.equals(""))
							                        	cell.setCellValue(included_parts);
							                        else
							                        	cell.setCellValue(" ");
							                        
							                        cell = row.createCell(TITLE_REVISION_DATE);
							                        if (revision_date != null && !revision_date.equals(""))
							                        	cell.setCellValue(revision_date);
							                        else
							                        	cell.setCellValue(" ");
							                        
							                        cell = row.createCell(TITLE_CONTAINS);
							                        if (contain_descrip != null && !contain_descrip.equals(""))
							                        	cell.setCellValue(contain_descrip);
							                        else
							                        	cell.setCellValue(" ");
							                        
							                        cell = row.createCell(TITLE_DATE_PUBLISHED);
							                        if (date_published != null && !date_published.equals(""))
							                        	cell.setCellValue(date_published);
							                        else
							                        	cell.setCellValue(" ");
							                        
							                        //SUBTITLE DETAILS
							                        cell = row.createCell(SUBTITLE_NAME_COLUMN);
							                        if (hd_source != null && !hd_source.equals(""))
							                        	cell.setCellValue(hd_source);
							                        else
							                        	cell.setCellValue(" ");
							                        
							                        //SUBPART DETAILS
							                        cell = row.createCell(SUBPART_COLUMN);
							                        if (header != null && !header.equals(""))
							                        	cell.setCellValue(header);
							                        else
							                        	cell.setCellValue(" ");
							                        	
													//SECTION DETAILS
							                        cell = row.createCell(SECTION_NUM_COLUMN);
							                        if (ss_num != null && !ss_num.equals(""))
							                        	cell.setCellValue(ss_num);
							                        else
							                        	cell.setCellValue(" ");
							                        	
							                        cell = row.createCell(SECTION_NAME_COLUMN);
							                        if (ss_subject != null && !ss_subject.equals(""))
							                        	cell.setCellValue(ss_subject);
							                        else
							                        	cell.setCellValue(" ");
							
							                        cell = row.createCell(SUBREQ_PARA_COLUMN);
							                        if (ss_para != null && !ss_para.equals(""))
							                        	cell.setCellValue(ss_para);
							                        else
							                        	cell.setCellValue(" ");
							
							                        cell = row.createCell(CITA_COLUMN);
							                        if (citation != null && !citation.isEmpty())
							                        	cell.setCellValue(citation);
							                        else
							                        	cell.setCellValue(" ");
					                        	}
					                        }
					                    }
					                }
					            }
					        }
				    	}
					}
				}
			}
		//}
	//}
        // File fileResult = new File("C:/Desktop/example.xlsx");
        // FileOutputStream fileOut = new FileOutputStream(fileResult);
        // create the file if it doesn't exist
        // if (!fileResult.exists()){
        // 	fileResult.createNewFile();
        // }
        
        FileOutputStream fileOut = new FileOutputStream("example.xlsx");
        workbook.write(fileOut);
        workbook.close();
        fileOut.close();

        /* 
        //To delete XML file
        
        if (xmlFile.exists()) {
            System.out.println("delete file-> " + xmlFile.getAbsolutePath());
            if (!xmlFile.delete()) {
                System.out.println("file '" + xmlFile.getAbsolutePath() + "' was not deleted!");
            }
        }
		*/
		
        System.out.println("Read of XML is COMPLETE.");
    }


    /**
     * Initializes the POI workbook and writes the header row
     */
    private static void initXls() {
        workbook = new XSSFWorkbook();

        CellStyle style = workbook.createCellStyle();
        Font boldFont = workbook.createFont();
        boldFont.setBold(true);
        style.setFont(boldFont);
        style.setAlignment(CellStyle.ALIGN_CENTER);

        Sheet sheet = workbook.createSheet();
        rowNum = 0;
        Row row = sheet.createRow(rowNum++);
        
        Cell cell = row.createCell(TITLE_NUM_COLUMN);
        cell.setCellValue("TITLE");
        cell.setCellStyle(style);
		
		cell = row.createCell(TITLE_SUBJECT);
        cell.setCellValue("SUBJECT");
        cell.setCellStyle(style);
        
        cell = row.createCell(TITLE_PARTS);
        cell.setCellValue("SCOPE");
        cell.setCellStyle(style);
        
        cell = row.createCell(TITLE_REVISION_DATE);
        cell.setCellValue("LAST REVISION");
        cell.setCellStyle(style);
        
        cell = row.createCell(TITLE_CONTAINS);
        cell.setCellValue("DESCRIPTION");
        cell.setCellStyle(style);
        
        cell = row.createCell(TITLE_DATE_PUBLISHED);
        cell.setCellValue("DATE PUBLISHED");
        cell.setCellStyle(style);
                
        //SUBTITLE COLUMNS
        cell = row.createCell(SUBTITLE_NAME_COLUMN);
        cell.setCellValue("SUBTITLE");
        cell.setCellStyle(style);
        
        cell = row.createCell(CHAPTER_NUM_COLUMN);
        cell.setCellValue("CHAPTER");
        cell.setCellStyle(style);
        
        //PART COLUMNS
        cell = row.createCell(PART_NUM_COLUMN);
        cell.setCellValue("PART#");
        cell.setCellStyle(style);
        
        cell = row.createCell(PART_NAME_COLUMN);
        cell.setCellValue("PART");
        cell.setCellStyle(style);
        
        //SUBPART COLUMNS
		cell = row.createCell(SUBPART_COLUMN);
        cell.setCellValue("SUBPART");
        cell.setCellStyle(style);
        
        //SECTION COLUMNS
        cell = row.createCell(SECTION_NUM_COLUMN);
        cell.setCellValue("SECTION#");
        cell.setCellStyle(style);

        cell = row.createCell(SECTION_NAME_COLUMN);
        cell.setCellValue("SECTION");
        cell.setCellStyle(style);

        cell = row.createCell(SUBREQ_PARA_COLUMN);
        cell.setCellValue("SUB-REQUIREMENT");
        cell.setCellStyle(style);

        cell = row.createCell(CITA_COLUMN);
        cell.setCellValue("FR CITATION");
        cell.setCellStyle(style);
    }
}
