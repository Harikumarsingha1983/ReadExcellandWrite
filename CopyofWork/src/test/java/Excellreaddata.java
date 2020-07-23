

import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excellreaddata {
	static XSSFSheet sheet;
	String ProjectPath;
	static XSSFWorkbook workbook;

	public Excellreaddata(String excellpath, String sheetname )
	{
		try {

			workbook = new XSSFWorkbook(excellpath);
			sheet = workbook.getSheet(sheetname);
		} catch (IOException e) {

			e.printStackTrace();
		}

	}

	public static int getRowCount() {
		int rowCount=0;
		try {
			rowCount = sheet.getPhysicalNumberOfRows();
			System.out.println("Number of rows are "+rowCount);
		} catch (Exception e) {
			System.out.println(e.getCause());
			System.out.println(e.getMessage());
			e.printStackTrace();
		}return rowCount;

	}
	public static int getColCount() {
		int colCount=0;
		try{
			 colCount = sheet.getRow(0).getPhysicalNumberOfCells();
		System.out.println("Number of rows are "+colCount);
		}catch (Exception e) {
			System.out.println(e.getCause());
			System.out.println(e.getMessage());
			e.printStackTrace();
		}
		return colCount;
	}

	public static String getCelldatastring(int rowNum,int colNum )
	{
		String CellData=null;
		try {
			CellData= sheet.getRow(rowNum).getCell(colNum).getStringCellValue();
			System.out.println(CellData);
		} catch (Exception e) {
			System.out.println(e.getCause());
			System.out.println(e.getMessage());
			e.printStackTrace();
		}return CellData;
	}

	public static double getCelldataNumber(int rowNum,int colNum )
	{double celldata=0;
		try {
			 celldata = sheet.getRow(rowNum).getCell(colNum).getNumericCellValue();
		} catch (Exception e) {
			System.out.println(e.getCause());
			System.out.println(e.getMessage());
			e.printStackTrace();
		}
return celldata;

	}
	
	
}

