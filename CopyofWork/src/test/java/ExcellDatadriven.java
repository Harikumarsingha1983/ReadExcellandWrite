

import java.io.File;
import java.io.FileNotFoundException;
import java.io.PrintStream;

import org.openqa.selenium.remote.html5.AddLocationContext;

public class ExcellDatadriven {
	public static void main(String[] args) {
		String excellpath = "PATH OF FILE";
		testdata(excellpath,"Commerical  Weekly");
		
		//Below codes are not required
		/*try
		{
			PrintStream myconsole = new PrintStream(new File("D:\\ECLIPSE\\Testing_project\\NewProject\\Hariconsole.txt"));
			System.setOut(myconsole);
			//myconsole.print(testdata);
			
		}catch(FileNotFoundException e)
		{
			System.out.println(e);
		}*/
	}

	public static void testdata(String excellpath, String Sheetname)
	{
		Excellreaddata excell = new Excellreaddata(excellpath,Sheetname);

		int rowCount= excell.getRowCount();
		int colCount= excell.getColCount();

		for (int i=0;i<rowCount;i++)
		{
			/*for (int j=0;i<colCount;j++)
			{
			 */			

			//String LOB= excell.getCelldatastring(i,0);

			String Composite= excell.getCelldatastring(i,3);
			double Results = (excell.getCelldataNumber(i, 6)*100);

			System.out.println(Composite+"                           "+Results);


		}
	}
}

