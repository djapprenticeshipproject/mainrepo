package djapprenticeshipproject;



import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
 
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class msExcel {
	
	
	
	//constants
	private static final String EXCEL_FILE_LOCATION = "C:\\JavaWorkspace\\mainrepo\\Docs\\Virtual Machines.xlsx";
	
	//private methods
	private static void outputSheet(Workbook w) { 
        try {
        	FileOutputStream outputStream = new FileOutputStream("JavaTESTBooks.xlsx");
			w.write(outputStream);
			w.close();
			outputStream.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
    

	}

	//public methods
	public static void appendNewRow() throws InvalidFormatException
	{ //method to add a number of new rows to the spreadsheet. 
		try {
            FileInputStream inputStream = new FileInputStream(new File(EXCEL_FILE_LOCATION));
            Workbook workbook = WorkbookFactory.create(inputStream);
 
            Sheet sheet = workbook.getSheetAt(0);
 
            Object[][] VMData = {
                    {"VM151", "AnotherBot", "ABot Live","100086","fakepassword86","", "OE Robotics", "AVM1086","LiveCHS","RTRA",""},
                    {"VM152", "AnotherBot", "ABot Live","100087","fakepassword87","", "OE Robotics", "AVM1087","LiveCHS","RTRA",""},
                    {"VM153", "AnotherBot", "ABot Live","100088","fakepassword86","", "OE Robotics", "AVM1088","LiveCHS","RTRA",""},
                    {"VM154", "AnotherBot", "ABot Live","100089","fakepassword86","", "OE Robotics", "AVM1089","LiveCHS","RTRA",""},
            };
 
            //find the Y position of the last row in the sheet into rowCount
            int rowCount = sheet.getLastRowNum();
 
            //loop through the VMData above 
            for (Object[] aVM : VMData) {
                Row row = sheet.createRow(++rowCount);
 
                int columnCount = 0;
       
                //loop through each row in the 2D array "aVM" 
                for (Object field : aVM) {
                	Cell cell = row.createCell(columnCount);
                    if (field instanceof String) {
                    cell.setCellValue((String) field);
                    	}
                    //increment the columnCount to move to the next column before looping
                    columnCount = columnCount +1;
                	} 
            }
            inputStream.close();
            //call private method to output the workbook
            outputSheet(workbook);
            
        } catch (IOException | EncryptedDocumentException ex) {
            ex.printStackTrace();
        }
    }
	
	public static void appendInputRow(String[] newRow)  throws InvalidFormatException
	{
		//method to append an input row constructed in the string array passed into 
		try {
            FileInputStream inputStream = new FileInputStream(new File("JavaTESTBooks.xlsx"));
            Workbook workbook = WorkbookFactory.create(inputStream);

            Sheet sheet = workbook.getSheetAt(0);
            
            //find the Y position of the last row in the sheet into rowCount
            int rowCount = sheet.getLastRowNum();
            Row row = sheet.createRow(++rowCount);
            int columnCount = 0;

            //loop through the newRow Array adding each element to a new column position 
            for (String element : newRow) {
            	Cell cell = row.createCell(columnCount);
            	if (element instanceof String) {
                    cell.setCellValue((String) element);
                    	}
            	//increment the columnCount to move to the next column before looping
                columnCount = columnCount +1;
            };
 
            inputStream.close();
            //call private method to output the workbook
            outputSheet(workbook);
 
        } catch (IOException | EncryptedDocumentException ex) {
            ex.printStackTrace();
        }
    }

	
	
};

