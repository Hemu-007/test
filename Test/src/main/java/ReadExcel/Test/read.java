package ReadExcel.Test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.commons.math3.analysis.function.Ceil;
import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class read {

	public static void main(String[] args) throws Exception 
	{
	   //I have placed an excel file 'Test.xlsx' in my D Driver 
	   FileInputStream fis = new FileInputStream("C:\\Users\\Hemu\\OneDrive\\Desktop\\Book1.xlsx");
	   
	  // HSSFWorkbook wb=new HSSFWorkbook(inputStream);
	   XSSFWorkbook workbook = new XSSFWorkbook(fis);
	   XSSFSheet sheet = workbook.getSheetAt(0);
	   //I have added test data in the cell A1 as "SoftwareTestingMaterial.com"
	/*   //Cell A1 = row 0 and column 0. It reads first row as 0 and Column A as 0.
		Row row = sheet.getRow(0);
		Cell cell = row.getCell(0);
		//System.out.println(cell);
		System.out.println(sheet.getRow(0).getCell(0));
		System.out.println(sheet.getRow(1).getCell(0));
		System.out.println(sheet.getRow(2).getCell(0));
		String cellval = cell.getStringCellValue();
		System.out.println(cellval);
		*/
		
		XSSFSheet sheet1 = workbook.getSheetAt(0);
		int lastRow = sheet1.getLastRowNum();
		for(int i=0; i<=lastRow; i++)
		{
		Row row = sheet1.getRow(i);
		Cell cell = row.createCell(2);

		cell.setCellValue("pushpa na maga hemu");
		
		}
		
		
		FileOutputStream fos = new FileOutputStream("C:\\Users\\Hemu\\OneDrive\\Desktop\\Book2.xlsx");
		workbook.write(fos);
		fos.close();
					
					
		}
		
		
		

	}


