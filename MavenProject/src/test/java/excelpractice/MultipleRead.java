package excelpractice;

import java.io.FileInputStream;
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

public class MultipleRead {

	public static void main(String[] args) throws EncryptedDocumentException, IOException {

		FileInputStream fis=new FileInputStream("./src/test/resources/testData/Sheet1.xlsx");
		Workbook workbook= WorkbookFactory.create(fis);
		Sheet sheet1=workbook.getSheet("Sheet1");
		Sheet sheet2=workbook.getSheet("Sheet2");
		int rowsize1=sheet1.getLastRowNum()-sheet1.getFirstRowNum();
		//int colsize1= sheet1.getRow(0).getPhysicalNumberOfCells();
		for(int i =0;i<rowsize1+1;i++) {
			Row row1=sheet1.getRow(i);  //0
			Row row2=sheet2.getRow(i);
			for(int j=0;j<row1.getLastCellNum();j++ ) {
				Cell cell1=row1.getCell(j);
				Cell cell2=row2.getCell(j);
				String val1=cell1.getStringCellValue();
				String val2=cell2.getStringCellValue();
				/*System.out.println(val1);
				System.out.println(val2);*/
				if(val1.equalsIgnoreCase(val2)) {
					System.out.println("it is equal");

				}if(val1.equalsIgnoreCase(val2)) {
					System.out.println("it is not equal");
					CellStyle cellStyle= workbook.createCellStyle();
					System.out.println(cellStyle);
					cellStyle.setFillBackgroundColor(IndexedColors.AQUA.getIndex());
					cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
					cellStyle.setFillForegroundColor(IndexedColors.BLUE.getIndex());
					cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
					cell2.setCellStyle(cellStyle);
				}
			}
		}
	}
	

}
