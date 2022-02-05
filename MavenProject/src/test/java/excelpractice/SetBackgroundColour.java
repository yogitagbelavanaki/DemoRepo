package excelpractice;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class SetBackgroundColour {
	
	public static void main(String[] args) throws EncryptedDocumentException, IOException {
		
		FileInputStream fis=new FileInputStream("./src/test/resources/testData/Sheet.xlsx");
		Workbook workbook= WorkbookFactory.create(fis);
		
		Sheet sheet2=workbook.getSheet("Sheet1");
		
		/*CellStyle backgroundStyle = workbook.createCellStyle();
		System.out.println(backgroundStyle);
		/*if(backgroundStyle==null) {
			System.out.println("null");
			
		}else {
			System.out.println("not null");
			backgroundStyle.setFillForegroundColor(IndexedColors.BLUE_GREY.getIndex());
			backgroundStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			Cell cell;
			cell.setCellStyle(backgroundStyle);
			*/
		Row row=sheet2.createRow(0);
		Cell cell= row.createCell((short)1);
		cell.setCellValue("Value");
		
		CellStyle cellstyle=workbook.createCellStyle();
		cellstyle.setFillBackgroundColor(IndexedColors.AQUA.getIndex());
		cellstyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		cell.setCellStyle(cellstyle);
		Cell cell1=row.createCell(2);
		cellstyle.setFillForegroundColor(IndexedColors.BLUE.getIndex());
		cellstyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		cell1.setCellStyle(cellstyle);
		FileOutputStream fos= new FileOutputStream("./src/test/resources/testData/Sheet.xlsx");
		workbook.write(fos);
		
		
	}

}
