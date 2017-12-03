import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;

//import java.io.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

public class Excel_Reader {

	public Excel_Reader() {
		// TODO Auto-generated constructor stub
	}

	public static void main(String[] args) {
		
        //accept file name or directory name through command line args
        String fname = args[0];
        
        //pass the filename or directory name to File object
        File f = new File(fname);
        
        if (f.exists()) 
        {
        	if (!f.canRead())
        	{
        		System.out.print("The File cannot be read!!!");
        		System.out.print("Check the file permissions for read!!!");
        	}
        } else
        {
        	System.out.printf("The file " + fname + " does not exist");
        	System.exit(0);
        }
	    try {
	        //FileInputStream file = new FileInputStream(new File("/home/sachin/Desktop/howtodoinjava_demo.xlsx"));
	        FileInputStream file = new FileInputStream(f);

	        //Create Workbook instance holding reference to .xlsx file
	        XSSFWorkbook workbook = new XSSFWorkbook(file);

	        //Get first/desired sheet from the workbook
	        XSSFSheet sheet = workbook.getSheetAt(0);

	        //Iterate through each rows one by one
	        Iterator<Row> rowIterator = sheet.iterator();
	        while (rowIterator.hasNext())
	        {
	            Row row = rowIterator.next();
	            //For each row, iterate through all the columns
	            Iterator<Cell> cellIterator = row.cellIterator();

	            while (cellIterator.hasNext()) 
	            {
	                Cell cell = cellIterator.next();
	                //Check the cell type and format accordingly
	                switch (cell.getCellType()) 
	                {
	                    case Cell.CELL_TYPE_NUMERIC:
	                        System.out.print(cell.getNumericCellValue() + "\t");
	                        break;
	                    case Cell.CELL_TYPE_STRING:
	                        System.out.print(cell.getStringCellValue() + "\t");
	                        break;
	                }
	            }
	            System.out.println("");
	        }
	        file.close();
	    } catch (Exception e) {
	        e.printStackTrace();
	    }
	}
}
