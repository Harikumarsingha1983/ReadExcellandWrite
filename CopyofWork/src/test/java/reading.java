import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class reading {
	
	public static void main(String[] args) throws IOException {
		try {
			FileInputStream fis = new FileInputStream("FILEPATH");				
			XSSFWorkbook book = new XSSFWorkbook(fis);
		      XSSFSheet sheet = book.getSheetAt(0);
		      
		      int rows = sheet.getLastRowNum();
		      
		      for (int i=0; i<=rows; i++)
		      {
		    	 Row row=sheet.getRow(i);
		    	  int cells = row.getLastCellNum();
		    	  for (int j=0;j<cells; j++) 
		    	  {
		    		  Cell cell =row.getCell(j);
		    		  if (cell.getCellType().name().equals("NUMERIC"))
		    		  {
		    			  System.out.println(cell.getNumericCellValue()+"\t");
		    		  }
		    		  else if (cell.getCellType().name().equals("STRING"))
		    		  {
		    			  System.out.println(cell.getStringCellValue()+"\t");
		    		  }
		    	  }System.out.println(sheet.getRowSumsBelow());
		    	  
		      }
			
		} catch (FileNotFoundException e) {
			System.out.println("File was not available or not in proper format");
			
			e.printStackTrace();
		}
	}


}
