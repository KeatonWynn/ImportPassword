package globalclasses;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.security.GeneralSecurityException;
import java.util.ArrayList;

import org.apache.poi.hssf.record.crypto.Biff8EncryptionKey;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


/**
 * 
 * The ExcelImport class is used to remove redundant code for excel imports. Currently, only .xlsx formats are able to be 
 * imported but will add .xls and .csv later when time permits. Only the first sheet will be loaded. The code will dynamically 
 * set the array sizes so it should be able to handle any amount of data given that you have required memory available 
 * 
 * 
 * 
 * @author HMF05046
 *
 */
public class ExcelImport {
	
	
	private String fl; 
	private String sn;
	private String pw;
	
	private int noc; //used to set number of columns to export
	private int sar; //indicates which row to stat reading from
	private boolean ColumnsSet = false;
	
	public ArrayList<String> ColumnNamesList = new ArrayList<String>();// Creating generic integer ArrayList
	public String ImportArray[][];
	public String ExportArray[][];
	public int RowCount;
	public int ColumnCount; 
	
	
	public void importXLSX(String FileLocation) throws IOException{
		
		fl = FileLocation;
		
		
				
		File myFile = new File(fl); 
		
		if(!myFile.exists()){
			System.out.println("File not found. Please verify that the file exists: " + fl);
			System.out.println("This program will terminate in 30 seconds.");
			System.exit(0);
		}
		
		FileInputStream fis = new FileInputStream(myFile); 
	    // Finds the workbook instance for xlsx file 
	    XSSFWorkbook myWorkBook = new XSSFWorkbook (fis);
	    // Return first sheet from the xlsx workbook 
	    XSSFSheet mySheet = myWorkBook.getSheetAt(0);
	    	    
	   
	    
	    
        // Get iterator to all the rows in current sheet        
        
        int NumberOfRows = mySheet.getLastRowNum();
        int NumberOfCol = mySheet.getRow(0).getPhysicalNumberOfCells();
        
        
        RowCount = NumberOfRows;
        ColumnCount = NumberOfCol;
        
               
        //Set size of Array
	    ImportArray = new String[NumberOfRows][NumberOfCol];
       		 	       	
        for(int i = 0; i < NumberOfRows; i++){
            
            //Get 2nd Row and then increases by one until reaching last row
            Row row = mySheet.getRow(1+i);                       
            
            //Dynamically creates Variable in Array, ImportArray[0][0] = cell "A2", ImportArray[0][1] = cell "A3", etc. 
            for(int j = 0; j <  NumberOfCol; j++){
            	ImportArray[0+i] [j]  = String.valueOf(row.getCell(j));
            }

            
        }
        
	}
	
public void importProtectedXLSX(String FileLocation, String Password) throws IOException, GeneralSecurityException{
		
		fl = FileLocation;
		pw = Password; 
		
		File myFile = new File(fl); 
		
		if(!myFile.exists()){
			System.out.println("File not found. Please verify that the file exists: " + fl);
			System.out.println("This program will terminate in 30 seconds.");
			System.exit(0);
		}
		
		NPOIFSFileSystem fileSystem = new NPOIFSFileSystem(new File(fl), true);
		EncryptionInfo info = new EncryptionInfo(fileSystem);
		Decryptor decryptor = Decryptor.getInstance(info);
		if (!decryptor.verifyPassword(fl)) {
			throw new RuntimeException("Unable to process: document is encrypted.");
		}
		InputStream dataStream = decryptor.getDataStream(fileSystem);
		Workbook myWorkBook = new XSSFWorkbook(dataStream);
		
			    // Return first sheet from the xlsx workbook 
	    Sheet mySheet = myWorkBook.getSheetAt(0);
	    	    
	   
	    
	    
        // Get iterator to all the rows in current sheet        
        
        int NumberOfRows = mySheet.getLastRowNum();
        int NumberOfCol = mySheet.getRow(0).getPhysicalNumberOfCells();
        
        
        RowCount = NumberOfRows;
        ColumnCount = NumberOfCol;
        
               
        //Set size of Array
	    ImportArray = new String[NumberOfRows][NumberOfCol];
       		 	       	
        for(int i = 0; i < NumberOfRows; i++){
            
            //Get 2nd Row and then increases by one until reaching last row
            Row row = mySheet.getRow(1+i);                       
            
            //Dynamically creates Variable in Array, ImportArray[0][0] = cell "A2", ImportArray[0][1] = cell "A3", etc. 
            for(int j = 0; j <  NumberOfCol; j++){
            	ImportArray[0+i] [j]  = String.valueOf(row.getCell(j));
            }

            
        }
        
	}
	
	//Export Methods
			
	public void setExportColumnNames(int ColumnNumber, String ColumnName){
		//must use before exportXLSX method
		ColumnNamesList.add(ColumnNumber, ColumnName); 
		
		if (ColumnNamesList.size() > 0){
			ColumnsSet = true; 
		}
		
		ColumnCount = ColumnNamesList.size(); 
	}

		
	public void exportXLSX(String FileLocation, String SheetName) throws IOException{
		//must use setExportColumnNames before this method
		
		if (ColumnsSet == true){
			sn = SheetName;

			//create new workbook
			XSSFWorkbook wb = new XSSFWorkbook();	       

			//create new worksheet
			XSSFSheet mySheet = wb.createSheet(sn);
			
			int NumberOfRows = ExportArray.length;
	        int NumberOfCol = ColumnNamesList.size();
	        RowCount = NumberOfRows;
	        ColumnCount = NumberOfCol;
	        
	        System.out.println("Number of rows created: " + NumberOfRows);
	        System.out.println("Number of columns created: " + NumberOfCol);
	       
      
			// Create a row and put some cells in it. Rows are 0 based. Creates Column Header
			Row row_column_header = mySheet.createRow((short)0);


			// will dynamically create column headers using values set in setExportColumnNames method
			for(int i=0; i < NumberOfCol; i++){
				row_column_header.createCell(i).setCellValue(ColumnNamesList.get(i));
				System.out.println("Created Column " + i + " = " + ColumnNamesList.get(i));
			}
			
			for(int i = 0; i < NumberOfRows; i++){
	            
				Row row = mySheet.createRow((short)1+i);                  
	            
	            //Dynamically creates Variable in Array, ImportArray[0][0] = cell "A2", ImportArray[0][1] = cell "A3", etc. 
	            for(int j = 0; j <  NumberOfCol; j++){
	            	row.createCell(j).setCellValue(ExportArray[i][j]);
	            }
	        }
			
			FileOutputStream fileOut = new FileOutputStream(FileLocation);
			wb.write(fileOut);
			wb.close(); 
		    fileOut.close();
			
		}
		else{
			System.out.println("Export columns not set. Please use setExportColumnNames method to set columns and try again.");
		}
	}
	
	//mainly used to copy files that need rows removed 
	public void copyXLSX(String ImportFile, String Exportfile, String ExportSheetName, int StartAtRow) throws IOException{
		
		sar = StartAtRow;
		sn = ExportSheetName;
		
		File myFile = new File(ImportFile); 
		
		if(!myFile.exists()){
			System.out.println("File not found. Please verify that the file exists: " + fl);
			System.out.println("This program will terminate in 30 seconds.");
			System.exit(0);
		}
		
		FileInputStream fis = new FileInputStream(myFile); 
	    // Finds the workbook instance for xlsx file 
	    XSSFWorkbook myWorkBook = new XSSFWorkbook (fis);
	    // Return first sheet from the xlsx workbook 
	    XSSFSheet mySheet = myWorkBook.getSheetAt(0);
	    	    
	   
        // Get iterator to all the rows in current sheet        
        
        int NumberOfRows = mySheet.getLastRowNum()-sar;
        int NumberOfCol = mySheet.getRow(sar).getPhysicalNumberOfCells();
        
        
        RowCount = NumberOfRows;
        ColumnCount = NumberOfCol;
        
               
        //Set size of Array
	    ImportArray = new String[NumberOfRows][NumberOfCol];
	    ExportArray = new String[NumberOfRows][NumberOfCol];
       		 	       	
        for(int i = 0; i < NumberOfRows; i++){
            
            //Get 2nd Row and then increases by one until reaching last row
            Row row = mySheet.getRow(1+i);                       
            
            //Dynamically creates Variable in Array, ImportArray[0][0] = cell "A2", ImportArray[0][1] = cell "A3", etc. 
            for(int j = 0; j <  NumberOfCol; j++){
            	ImportArray[0+i] [j]  = String.valueOf(row.getCell(j));
            }

            
        }
        
        
        //create new workbook
		XSSFWorkbook wb = new XSSFWorkbook();	       

		//create new worksheet
		XSSFSheet ExportmySheet = wb.createSheet(sn);
		        
        System.out.println("Number of rows created: " + NumberOfRows);
        System.out.println("Number of columns created: " + NumberOfCol);
       
  
		// Create a row and put some cells in it. Rows are 0 based. Creates Column Header
		Row row_column_header = ExportmySheet.createRow((short)0);


		// will dynamically create column headers using values set in setExportColumnNames method
		for(int i=0; i < NumberOfCol; i++){
			row_column_header.createCell(i).setCellValue(ImportArray[0] [0+i]);
			System.out.println("Created Column " + i + " = " + ImportArray[0] [0+i]);
		}
		
		for(int i = 0; i < NumberOfRows; i++){
            
			Row row = ExportmySheet.createRow((short)1+i);                  
            
            //Dynamically creates Variable in Array, ImportArray[0][0] = cell "A2", ImportArray[0][1] = cell "A3", etc. 
            for(int j = 0; j <  NumberOfCol; j++){
            	row.createCell(j).setCellValue(ImportArray[i][j]);
            }
        }
		
		FileOutputStream fileOut = new FileOutputStream(Exportfile);
		wb.write(fileOut);
		wb.close(); 
	    fileOut.close();
		
		
	}
	
public void copyXLS(String ImportFile, String Exportfile, String ExportSheetName, int StartAtRow) throws IOException{
		
		sar = StartAtRow;
		sn = ExportSheetName;
		
		File myFile = new File(ImportFile); 
		
		if(!myFile.exists()){
			System.out.println("File not found. Please verify that the file exists: " + fl);
			System.out.println("This program will terminate in 30 seconds.");
			System.exit(0);
		}
		
		FileInputStream fis = new FileInputStream(myFile); 
	    // Finds the workbook instance for xlsx file 
		HSSFWorkbook myWorkBook = new HSSFWorkbook (fis);
	    // Return first sheet from the xlsx workbook 
	    HSSFSheet mySheet = myWorkBook.getSheetAt(0);
	    	    
	   
        // Get iterator to all the rows in current sheet        
        
	    int NumberOfRows = mySheet.getLastRowNum()-sar;
        int NumberOfCol = mySheet.getRow(sar).getPhysicalNumberOfCells();
        
        
        RowCount = NumberOfRows;
        ColumnCount = NumberOfCol;
        
               
        //Set size of Array
	    ImportArray = new String[NumberOfRows][NumberOfCol];
	    ExportArray = new String[NumberOfRows][NumberOfCol];
       		 	       	
        for(int i = 0; i < NumberOfRows; i++){
            
            //Get 2nd Row and then increases by one until reaching last row
            Row row = mySheet.getRow(sar+i);                       
            
            //Dynamically creates Variable in Array, ImportArray[0][0] = cell "A2", ImportArray[0][1] = cell "A3", etc. 
            for(int j = 0; j <  NumberOfCol; j++){
            	ImportArray[0+i] [j]  = String.valueOf(row.getCell(j));
            }
            
            
        }
        
        
        //create new workbook
        HSSFWorkbook wb = new HSSFWorkbook();	       

		//create new worksheet
		HSSFSheet ExportmySheet = wb.createSheet(sn);
		        
        System.out.println("Number of rows created: " + NumberOfRows);
        System.out.println("Number of columns created: " + NumberOfCol);
       
  
		// Create a row and put some cells in it. Rows are 0 based. Creates Column Header
		Row row_column_header = ExportmySheet.createRow((short)0);


		// will dynamically create column headers using values set in setExportColumnNames method
		for(int i=0; i < NumberOfCol; i++){
			row_column_header.createCell(i).setCellValue(ImportArray[0] [0+i]);
			System.out.println("Created Column " + i + " = " + ImportArray[0] [0+i]);
		}
		
		for(int i = 0; i < NumberOfRows; i++){
            
			Row row = ExportmySheet.createRow((short)0+i);                  
            
            //Dynamically creates Variable in Array, ImportArray[0][0] = cell "A2", ImportArray[0][1] = cell "A3", etc. 
            for(int j = 0; j <  NumberOfCol; j++){
            	row.createCell(j).setCellValue(ImportArray[i][j]);
            }
        }
		
		FileOutputStream fileOut = new FileOutputStream(Exportfile);
		wb.write(fileOut);
		wb.close(); 
	    fileOut.close();
		
		
	}
		
}
	
	
	
	

