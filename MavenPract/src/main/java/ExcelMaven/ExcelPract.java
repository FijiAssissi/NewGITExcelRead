package ExcelMaven;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelPract 
{
	XSSFSheet sh;//XSSFsheet is a class in apache poi- declaring here so it can be accessed by another methods inside the class
	public ExcelPract() throws IOException //constructor loads the excel file
	{
		//open excel file
		FileInputStream f= new FileInputStream("C:\\Users\\jerri\\OneDrive\\Documents\\JAVA Programs\\ExcelFiles\\ExcelMavenNewRead.xlsx");
		
		//load the workbook
		XSSFWorkbook wb= new XSSFWorkbook(f);//without passing f , it creates a new blank workbook, instead of loading an existing one
		//XSSFWorkbook is used because the file format is .xlsx.
		//We pass f (FileInputStream) into XSSFWorkbook to tell Java which Excel file to load.
		//get the sheet
		sh=wb.getSheet("Sheet1");//retrieves the existing sheet "Sheet 1" from excel file and storing in the variable
		
		//now the sheet is ready for reading
	}
		public String readData (int i,int j)// return the data from the specified cell as a String.
		{
			Row r=sh.getRow(i);//for the first row//get the row at the index i
			Cell c=r.getCell(j);//for the first column//get the column at the index j
			//the cell stored in the variable c
			
			//determine the cell type
			
			int celltype=c.getCellType();//getCellType  which returns an integer 
			//Retrieves the type of data stored in the cell 
			
			switch (celltype)//switch(expression)
			{
			case Cell.CELL_TYPE_NUMERIC://CELL_TYPE_NUMERIC is a constant defined inside the Cell interface so we need class/interface which holds them
				
				double d=c.getNumericCellValue();//retrieves the numeric value 
				return String.valueOf(d);//converts to string using String.valuesOf()
				
			case Cell.CELL_TYPE_STRING:
				String s=c.getStringCellValue();//retrieves it 
				return s;//no need to convert it to String, simply return 
			}
		
			return null;
		}
		
		
	
	public static void main(String[] args) 
	{
		
	}

}//
