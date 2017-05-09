package com.sowmindra;

import java.io.File;
import java.io.FileOutputStream;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.w3c.dom.Element;


public class XMLToExcelTest {

    private static Workbook workbook;
    private static Sheet sheet;
    private static int rowNum;
    private static DocumentBuilderFactory dbFactory;
    private static DocumentBuilder dBuilder;
    private static int folderNum;
    private static int fileNum;

    public static void main(String[] args) throws Exception {
    	String rootPath = "//10.96.45.13/c$/Program Files/OpenText/Process Platform/defaultInst/cws/sync/system/PeopleCenter";
    	//String rootPath = "//10.96.45.12/c$/Program Files/OpenText/Process Platform/defaultInst/cws/sync/peoplecenter/People Center";
    	//String rootPath = "C:/Sowmindra/junk/Hiring Management/TranslationInformation_00505601-1849-11E7-E128-5AE47856D0D0#cws-ma#.cws";
    	
    	System.out.println("Input " + rootPath);
    	
    	File rootDirectory = new File (rootPath);
    	
    	if(!rootDirectory.isDirectory())
    		System.out.println("Not a directory/folder");

        dbFactory = DocumentBuilderFactory.newInstance();
        dBuilder = dbFactory.newDocumentBuilder();

		initXls();
		sheet = workbook.getSheetAt(0);
		
		folderNum = -1;
		fileNum = 0;
		
       	traverseDirectory(rootDirectory);
        
        dbFactory = null;
        dBuilder = null;

        System.out.println("\nDone parsing XMLs");
        FileOutputStream fileOut = new FileOutputStream("C:/Temp/Translation-List.xlsx");
        workbook.write(fileOut);
        workbook.close();
        fileOut.close();
        System.out.println("Finished with Excel file as well");
    }

    private static void traverseDirectory(File directory) {
        
		File[] fileList = directory.listFiles();
		
		System.out.println("\nParsing folder " + directory.getName());
		
		fileNum = 0;
		folderNum++;
		
		for(int j=0; j<fileList.length;j++) {
	        if(fileList[j].isDirectory())
	        	traverseDirectory(fileList[j]);
	        else
	        	processXMLFile(fileList[j]);
		}
    }

    private static void processXMLFile(File xmlFile)  {
        String elementNames[] = {"DocumentID", "Name", "Description", "CreatedBy", "CreationDate", "Notes", "DefaultText"};

        Cell cell = null;

        int colCount = -1;
        
        fileNum++;
 
		System.out.print("\nFolder "+folderNum+"-"+fileNum+"-->Processing file " + xmlFile.getName());
		
		if(!xmlFile.getName().endsWith(".xml") && !xmlFile.getName().endsWith(".cws")) {
			System.out.print(" ^skipping");
			return;
		}
		
		try {
			
			Document doc = dBuilder.parse(xmlFile);
			if(doc.getElementsByTagName("TextIdentifier").getLength()==0 || doc.getElementsByTagName("TextIdentifier").getLength()>1)
				throw new Exception("File mismatch");
			
			Row row = sheet.createRow(rowNum++);
			
			cell = row.createCell(++colCount);
			cell.setCellValue(xmlFile.getName());
			
			cell = row.createCell(++colCount);
			cell.setCellValue(((Element)doc.getElementsByTagName("TextIdentifier").item(0))==null?"":((Element)doc.getElementsByTagName("TextIdentifier").item(0)).getAttribute("typeVersion"));
			
			cell = row.createCell(++colCount);
			cell.setCellValue(((Element)doc.getElementsByTagName("TextIdentifier").item(0))==null?"":((Element)doc.getElementsByTagName("TextIdentifier").item(0)).getAttribute("RuntimeDocumentID"));
			
			for(int i=0;i<elementNames.length;i++){
				cell = row.createCell(++colCount);
			    cell.setCellValue(((Element)doc.getElementsByTagName(elementNames[i]).item(0))==null?"":((Element)doc.getElementsByTagName(elementNames[i]).item(0)).getTextContent());
			}
			
			for(int i=0;i<doc.getElementsByTagName("uri").getLength();i++){	        	
				cell = row.createCell(++colCount);
				cell.setCellValue(((Element)doc.getElementsByTagName("uri").item(i)).getAttribute("id"));
			}
			
			cell = row.createCell(++colCount);
			cell.setCellValue(xmlFile.getAbsolutePath());

			System.out.print(" done");
		}
		catch (Exception e) {
			System.out.print(" **error " + e.getMessage());
		}
    }

    private static void initXls() {
        workbook = new XSSFWorkbook();

        CellStyle style = workbook.createCellStyle();
        Font boldFont = workbook.createFont();
        boldFont.setBold(true);
        style.setFont(boldFont);
        style.setAlignment(CellStyle.ALIGN_CENTER);

        Sheet sheet = workbook.createSheet("Translations");
        rowNum = 0;
        int colCount = 0;
        Row row = sheet.createRow(rowNum++);
        Cell cell = null; 
        
        cell = row.createCell(colCount++);
        cell.setCellValue("File Name");
        cell.setCellStyle(style);

        cell = row.createCell(colCount++);
        cell.setCellValue("typeVersion");
        cell.setCellStyle(style);

        cell = row.createCell(colCount++);
        cell.setCellValue("RuntimeDocumentID");
        cell.setCellStyle(style);

        cell = row.createCell(colCount++);
        cell.setCellValue("DocumentID");
        cell.setCellStyle(style);

        cell = row.createCell(colCount++);
        cell.setCellValue("Name");
        cell.setCellStyle(style);

        cell = row.createCell(colCount++);
        cell.setCellValue("Description");
        cell.setCellStyle(style);

        cell = row.createCell(colCount++);
        cell.setCellValue("CreatedBy");
        cell.setCellStyle(style);

        cell = row.createCell(colCount++);
        cell.setCellValue("CreationDate");
        cell.setCellStyle(style);

        cell = row.createCell(colCount++);
        cell.setCellValue("Notes");
        cell.setCellStyle(style);

        cell = row.createCell(colCount++);
        cell.setCellValue("DefaultText");
        cell.setCellStyle(style);

        cell = row.createCell(colCount++);
        cell.setCellValue("TranslationInformation-uri-id");
        cell.setCellStyle(style);

        cell = row.createCell(colCount++);
        cell.setCellValue("Translations-uri.id");
        cell.setCellStyle(style);

        cell = row.createCell(colCount++);
        cell.setCellValue("Full path");
        cell.setCellStyle(style);
     }
}
