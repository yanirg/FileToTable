import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.util.Iterator;
import java.util.List;
import java.util.Scanner;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

public class FileToWordTable {
	 FileSplitter fs;
	 Scanner sc = null;
	XWPFDocument document = new XWPFDocument();
    
    //create table
    XWPFTable table = document.createTable();
    XWPFTableRow tableRow = table.getRow(0);
    FileOutputStream out =null;
    int columnSize=0;

	public FileToWordTable( String outputFile,List<String> columnsName) {
		try {
			this.out =new FileOutputStream(new File(outputFile));
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} 
		initFirstRow(columnsName);
	}
	public FileToWordTable(String outputFile) {
		try {
			this.out =new FileOutputStream(new File(outputFile));
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} 
	}
	
	
	
	public void splitFileToTabale(Scanner fileToScan,String delemitrr) {
		
		while(fileToScan.hasNextLine()) {
			
			String loc = fileToScan.nextLine();
			String[] st = loc.split(delemitrr);
			//sc.skip(sc.delimiter());	
			//String str = sc.nextLine();
			tableRow = table.createRow();
			addRowToTable(tableRow,st);
		}
			 
	}
		
	private void addRowToTable(XWPFTableRow tableRow,String[] val) {
		for(int i=0;i<=columnSize;i++) {
			  tableRow.getCell(i).setText(val[i]);
		}
			
	}

    private void initFirstRow(List<String> columnsName) {
    	Iterator<String> it = columnsName.iterator();
    	columnSize=columnsName.size();
		int i =0;
		while(it.hasNext()) {
			if (i==0) {
				tableRow.getCell(0).setText(it.next());
				i++;
			}else {
				tableRow.addNewTableCell().setText(it.next() );
			}
		}
    }
    

}
