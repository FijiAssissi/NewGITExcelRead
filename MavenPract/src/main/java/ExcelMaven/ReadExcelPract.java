package ExcelMaven;

import java.io.IOException;

public class ReadExcelPract {

	public static void main(String[] args) throws IOException 
	{
	ExcelPract e= new ExcelPract();//calls the constructor, which loads excel file
	String s=e.readData(0, 0);//fetches data at row 0, column 0
	System.out.println(s);
		
	String p=e.readData(0, 1);
	System.out.println(p);
	}

}
