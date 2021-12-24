package Practice1;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

public class ReadExcel {
	
	@Test
	public void readexcel() throws IOException {
		//Reading Excel Data
		FileInputStream file = new FileInputStream("D:\\Practicesheet.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		XSSFSheet sheet = workbook.getSheet("Sheet1");
		XSSFSheet sheet2 = workbook.getSheetAt(1);
		
		System.out.println(sheet.getRow(0).getCell(0).getNumericCellValue());
		System.out.println(sheet.getRow(1).getCell(0).getStringCellValue());
		
		//Writing Data in Excel
		Row row = sheet2.createRow(1);
		Cell cell = row.createCell(1);
		cell.setCellValue("Vaibhav Gunjkar");
		FileOutputStream fos = new FileOutputStream("D:\\Practicesheet.xlsx");
		workbook.write(fos);
		System.out.println("END OF WRITING EXCEL");
		
		
		
	}

}
