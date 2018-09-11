
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.Scanner;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.hssf.usermodel.*;

public class ReadingExcel 
{
	
	public static void main(String[] args) throws IOException 
	{
		try
		{
			File excelFile = new File("C://Users/kramakri/Desktop/sample.xls");
			FileInputStream fis = new FileInputStream(excelFile);
	
			HSSFWorkbook workbook = new HSSFWorkbook(fis);
			HSSFSheet sheet = workbook.getSheet("Test1");
			HSSFRow rownum;
			HSSFCell cellValueFlag, value;
			Scanner input = new Scanner(System.in);
			for(int i=1; i <= sheet.getLastRowNum(); i++)
			{
				rownum = sheet.getRow(i);
				value = rownum.getCell(0);
				cellValueFlag = rownum.getCell(1);
				System.out.println("The value is" + cellValueFlag);
				if(cellValueFlag.getNumericCellValue() == 1 && cellValueFlag != null)
				{
					value = rownum.getCell(2);
					if(value == null)
					{
						value = rownum.createCell(2);
					}
					System.out.println("Enter the value for the flag" + value.toString());
					String userInput = input.nextLine();
					value.setCellValue(userInput);
				}
			}
	    
			fis.close();
			input.close();
 	    
			try(FileOutputStream fileOut = new FileOutputStream(excelFile))
			{
				workbook.write(fileOut);
			}
		}
		catch(FileNotFoundException e)
		{
			e.printStackTrace();
		}
	}
}
