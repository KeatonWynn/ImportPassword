package globalclasses;

import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Row;

/**
 * 
 * Current Website for each source (Last updated 09/10/2018): </br>
 * 
 * 	<li>AutoIMS = http://www.autoims.com/ </br>
 * 	<li>Adesa = https://www.adesa.com/home (?) </br>
 * 	<li>Manheim = https://mmr.manheim.com/ </br>
 * 	<li>RMA = http://irvhcablzs02v:7004/DeepRuleServiceRMA/ </br>
 * 	<li>OpenLane = https://www.adesa.com/home (?) </br>
 * 	<li>AutoIMSHCCA = http://www.autoims.com/ </br>
 * 
 * 
 * 
 * @author HMF05046
 *
 */
public class ImportPassword {
	
	private String user;
	private String pw;	
	private String userHMF = System.getProperty("user.name"); //Automatically gets user from system properties (i.e. HMF05046)
	private static String IA[][]; 
	
	public void importPassword() throws IOException {
		

		// Set file where passwords are stored on the hard drive.
		String myFile = "C:\\Users\\" + userHMF	+ "\\ImportFiles\\ImportFile.xlsx";

		// Create ExcelImport class and use file defaulf file location
		ExcelImport EI = new ExcelImport();		
		EI.importXLSX(myFile);
		//Another array (IA) needs to be created as the ImportArray cannot be used if it is non-static (hard to explain)
		IA = new String[EI.RowCount][EI.ColumnCount]; 
		for(int i=0; i < EI.RowCount ; i++){
			for(int j=0; j < EI.ColumnCount; j++){				
				IA[i][j] = EI.ImportArray[i][j]; 
			}
		}
	}

	public void importPassword(String FileLocation) throws IOException {

		// Set file where passwords are stored on the hard drive.
		String myFile = FileLocation;

		// Create ExcelImport class and use file defaul file location
		ExcelImport EI = new ExcelImport();		
		EI.importXLSX(myFile);
		//Another array (IA) needs to be created as the ImportArray cannot be used if it is non-static (hard to explain)
		IA = new String[EI.RowCount][EI.ColumnCount]; 
				for(int i=0; i < EI.RowCount ; i++){
					for(int j=0; j < EI.ColumnCount; j++){
						
						IA[i][j] = EI.ImportArray[i][j]; 
					}
				}

	}

	public String getUser(String Source) throws IOException,
			InterruptedException {

		int array_size = IA.length;

		for (int i = 0; i < array_size;) {

			if (Source.toUpperCase().equals(
					IA[0 + i][0].toUpperCase())) {

				user = IA[0 + i][1];
				pw = IA[0 + i][2];
//				System.out.println("Username and password found for " + Source	+ " source.");
				i = array_size;

			} else if (i == array_size - 1) {

				System.out.println("Username and password not found. Please check spelling or check excel file in location below:");
				System.out.println("C:\\Users\\" + userHMF	+ "\\Java_Files\\JavaPWImport.xlsx");
				i++;
			} else {
				i++;
				TimeUnit.SECONDS.sleep((long) 0.2);
			}
		}

		return user;

	}

	public String getPassword(String Source) throws IOException,
			InterruptedException {

		int array_size = IA.length;

		for (int i = 0; i < array_size;) {

			if (Source.toUpperCase().equals(
					IA[0 + i][0].toUpperCase())) {

				user = IA[0 + i][1];
				pw = IA[0 + i][2];
				i = array_size;

			} else if (i == array_size - 1) {

				System.out.println("Username and password not found. Please check spelling or check excel file in location below:");
				System.out.println("C:\\Users\\" + userHMF	+ "\\Java_Files\\JavaPWImport.xlsx");
				i++;
			} else {
				i++;
				TimeUnit.SECONDS.sleep((long) 0.2);
			}
		}

		return pw;

	}
	
	public String getPassword2(String Source) throws IOException,
	InterruptedException {
		
		int array_size = IA.length;

		for (int i = 0; i < array_size;) {

			if (Source.toUpperCase().equals(
					IA[0 + i][0].toUpperCase())) {

				user = IA[0 + i][1];
				pw = IA[0 + i][2];
				i = array_size;

			} else if (i == array_size - 1) {

				System.out.println("Username and password not found. Please check spelling or check excel file in location below:");
				System.out.println("C:\\Users\\" + userHMF	+ "\\Java_Files\\JavaPWImport.xlsx");
				i++;
			} else {
				i++;
				TimeUnit.SECONDS.sleep((long) 0.2);
			}
		}

		return pw;

	}
	
	
	
}
