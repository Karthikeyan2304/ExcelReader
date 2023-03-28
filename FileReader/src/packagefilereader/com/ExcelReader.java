package packagefilereader.com;

import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReader {

	public static void main(String[] args) throws IOException 
	{
	String Path=".\\ExcelFile\\StudentDetails.xlsx";
	FileInputStream fin=new FileInputStream(Path);
	XSSFWorkbook wb=new XSSFWorkbook(fin);
	XSSFSheet Sheet=wb.getSheetAt(0);
	
	
	
//	(for each Iterator) Iterator row
for (Row row : Sheet) 
{

	// (for each Iterator) Iterator Cell 	
	
	for (Cell cell : row)

		
		// To Check Cell Type I Used Switch 
		
	switch(cell.getCellType())
	{
	case STRING:
		System.out.print(cell.getStringCellValue()+"\t");break;
	case BOOLEAN:
		System.out.print(cell.getBooleanCellValue()+"\t");break;
	case NUMERIC:
		System.out.print(cell.getNumericCellValue()+"\t");break;
}System.out.println();
}
}
}	
