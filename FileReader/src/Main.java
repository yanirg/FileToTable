import java.io.BufferedReader;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.util.Arrays;
import java.util.Scanner;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

public class Main {
	static FileSplitter fs;
	static Scanner sc = null;
	 //Blank Document
   ;
    
    //Write the Document in file system
    
    //create table
   

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		//fs = new FileSplitter("C:\\Users\\yanir_000.NOTEBOOK-PC\\Desktop\\12.txt");
		XWPFDocument document = new XWPFDocument();

	      //Write the Document in file system
	      
	      //create table
	      XWPFTable table = document.createTable();

	      XWPFTableRow tableRow = table.getRow(0);
	      tableRow.getCell(0).setText("Column 1");

			tableRow.addNewTableCell().setText("Column 2");

			tableRow.addNewTableCell().setText("Column 3");
			FileOutputStream out =null;
	
		try {
		       out = new FileOutputStream(new File(System.getProperty("user.home") + "\\Desktop\\create_table.docx"));

			sc = new Scanner(new FileReader(System.getProperty("user.home") + "\\Desktop\\12.txt"));
			sc.useDelimiter(",");
			while(sc.hasNextLine()) {
				
				String loc = sc.nextLine();
				String[] st = loc.split(",");
				//sc.skip(sc.delimiter());	
				//String str = sc.nextLine();
				tableRow = table.createRow();
				addRowToTable(tableRow,st);
				 
			     

			    /*  //create second row
			      XWPFTableRow tableRowTwo = table.createRow();
			      
			      tableRowTwo.getCell(0).setText(st[1]);
			      tableRowTwo.getCell(1).setText("col two, row two");
			      tableRowTwo.getCell(2).setText("col three, row two");

			      //create third row
			      XWPFTableRow tableRowThree = table.createRow();
			      
			      tableRowThree.getCell(0).setText("col one, row three");
			      tableRowThree.getCell(1).setText("col two, row three");
			      tableRowThree.getCell(2).setText("col three, row three");
*/
			   
			      
				
				//System.out.println(Arrays.toString(st));
			}
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();		
	     
	   }
		   try {
				document.write(out);
				out.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		
		
	}
	
	
	public static void addRowToTable(XWPFTableRow tableRow,String[] val) {
		  tableRow.getCell(0).setText(val[0]);
	      tableRow.getCell(1).setText(val[1]);
	      tableRow.getCell(2).setText(val[2]);
		
	}
}

