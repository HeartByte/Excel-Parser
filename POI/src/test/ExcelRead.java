package test;

import java.io.FileInputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelRead {
	public static void main(String[] args) {

		try {
			FileInputStream file = new FileInputStream("./data/서울시CCTV정보.xls");

			XSSFWorkbook workbook = new XSSFWorkbook(file);			
			XSSFSheet sheet = workbook.getSheetAt(0);
			
			int r = 0;
			int c = 0;
			
			int rows = sheet.getPhysicalNumberOfRows();
			int cells = sheet.getRow(0).getPhysicalNumberOfCells();
			
			for (r = 0; r < rows; r++) {
                XSSFRow row = sheet.getRow(r); //  행 가져오기
                if (row != null) {
                    for (c = 0; c < cells; c++) {
                        XSSFCell cell = row.getCell(c);
                        if (cell != null) {
                            String value = "";
                            switch (cell.getCellType()) { // 다양한 형태의 엑셀 파일을 가져와서 집어 넣는다.
                            case FORMULA:
                                value = cell.getCellFormula();
                                break;
                            case NUMERIC:
                                value = "" + cell.getNumericCellValue();
                                break;
                            case STRING:
                                value = "" + cell.getStringCellValue();
                                break;
                            case BLANK:
                                value = "" + cell.getBooleanCellValue();
                                break;
                            case ERROR:
                                value = "" + cell.getErrorCellValue();
                                break;
                            default:
                            }
                           System.out.println(value + ",");  //Test Print
                          //  writeFile(value);
                        } 
                    } 
                   // writeFile("\n");
                    System.out.println(); //Test Print
                }
            } 
			
			
		} catch (Exception e) {
			e.printStackTrace();
		}
	
	}
}
