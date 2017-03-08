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
    private final static int SUBTITLE_NAME_COLUMN = 1;
    private final static int CHAPTER_NUM_COLUMN = 2;
    private final static int PART_NUM_COLUMN = 3;
    private final static int PART_NAME_COLUMN = 4;
    private final static int SUBPART_COLUMN = 5;
    private final static int SECTION_NUM_COLUMN = 6;
    private final static int SECTION_NAME_COLUMN = 7;
    private final static int SUBREQ_PARA_COLUMN = 8;
    private final static int CITA_COLUMN = 9;

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

        NodeList nList = doc.getElementsByTagName("SUBPART");
        for (int i = 0; i < nList.getLength(); i++) {
            System.out.println("Processing element " + (i + 1) + "/" + nList.getLength());
            Node node = nList.item(i);
            if (node.getNodeType() == Node.ELEMENT_NODE) {
                Element element = (Element) node;
                String header, section_num, subject;
                try {
                	header = element.getElementsByTagName("HD").item(0).getTextContent();
                	section_num = element.getElementsByTagName("SECTNO").item(0).getTextContent();
                	subject = element.getElementsByTagName("SUBJECT").item(0).getTextContent();
                } catch (NullPointerException e){
                	header = "";
                	section_num = "";
                	subject = "";
                }

                NodeList prods = element.getElementsByTagName("SECTION");
                for (int j = 0; j < prods.getLength(); j++) {
                	System.out.println("     Processing sub-element " + (j + 1) + "/" + prods.getLength());
                    Node prod = prods.item(j);
                    if (prod.getNodeType() == Node.ELEMENT_NODE) {
                        Element product = (Element) prod;
                        String ss_num, ss_subject, ss_para, citation;
                        try{
                        	ss_num = product.getElementsByTagName("SECTNO").item(0).getTextContent();
	                        ss_subject = product.getElementsByTagName("SUBJECT").item(0).getTextContent();
	                        ss_para = product.getElementsByTagName("P").item(0).getTextContent();
	                        citation = product.getElementsByTagName("CITA").item(0).getTextContent();
                        } catch (NullPointerException e){
                        	ss_num = "";
                        	ss_subject = "";
                        	ss_para = "";
                        	citation = "";
                        }

                        Row row = sheet.createRow(rowNum++);
                    
                        // check for NULL in .setCellValue()
                        Cell cell = row.createCell(SUBPART_COLUMN);
                        if (header != null && !header.equals(""))
                        	cell.setCellValue(header);
                        else
                        	cell.setCellValue(" Empty ");

                        cell = row.createCell(SECTION_NUM_COLUMN);
                        if (ss_num != null && !ss_num.equals(""))
                        	cell.setCellValue(ss_num);
                        else
                        	cell.setCellValue(" Empty ");

                        cell = row.createCell(SECTION_NAME_COLUMN);
                        if (ss_subject != null && !ss_subject.equals(""))
                        	cell.setCellValue(ss_subject);
                        else
                        	cell.setCellValue(" Empty ");

                        cell = row.createCell(SUBREQ_PARA_COLUMN);
                        if (ss_para != null && !ss_para.equals(""))
                        	cell.setCellValue(ss_para);
                        else
                        	cell.setCellValue(" Empty ");

                        //cell = row.createCell(CITA_COLUMN);
                        //if (citation != null && !citation.isEmpty())
                        //	cell.setCellValue(citation);
                        //else
                        //	cell.setCellValue(" Empty ");
                        
                    }
                }
            }
        }

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
		
        System.out.println("Read of XML is COMPLETE. \nPROCESSED " + nList.getLength() + " ELEMENTS");
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
        Cell cell = row.createCell(SUBPART_COLUMN);
        cell.setCellValue("Subpart");
        cell.setCellStyle(style);

        cell = row.createCell(SECTION_NUM_COLUMN);
        cell.setCellValue("Section Number");
        cell.setCellStyle(style);

        cell = row.createCell(SECTION_NAME_COLUMN);
        cell.setCellValue("Section Name");
        cell.setCellStyle(style);

        cell = row.createCell(SUBREQ_PARA_COLUMN);
        cell.setCellValue("Sub-Requirement");
        cell.setCellStyle(style);

        cell = row.createCell(CITA_COLUMN);
        cell.setCellValue("Citation");
        cell.setCellStyle(style);
    }
}
