import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



public class ReadandPulldata {
	
	
		   static XSSFRow row;
		  
		  public static void main(String[] args) throws IOException
			
		  {
		      FileInputStream fis = new FileInputStream(new File("FILEPATH"));
		      
		      XSSFWorkbook workbook = new XSSFWorkbook(fis);
		      XSSFSheet spreadsheet = workbook.getSheetAt(0);
		      Iterator < Row > rowIterator = spreadsheet.iterator();
		      
		      while (rowIterator.hasNext()) {
		         row = (XSSFRow) rowIterator.next();
		         Iterator < Cell >  cellIterator = row.cellIterator();
		         while ( cellIterator.hasNext()) {
		             Cell cell = cellIterator.next();
		             
		             switch (cell.getCellType()) 
		             {
		                case NUMERIC:
		                   System.out.print(cell.getNumericCellValue() + " \t\t ");
		                   break;
		                
		                case STRING:
		                   System.out.print(cell.getStringCellValue() + " \t\t ");
		                   break;
		                
		             }
		             /*XSSFCellStyle style1 = workbook.createCellStyle();
		             spreadsheet.setColumnWidth(0, 500);
		             style1.setAlignment(HorizontalAlignment.LEFT);           
		             cell.setCellValue("Top Left");
		             cell.setCellStyle(style1);  */ 
		            
		                     
		             
		          }
		         System.out.println(row.getFirstCellNum());
		         
		       }
		       fis.close();		    }
		  }  
		  
			         


