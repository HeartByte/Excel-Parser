package test;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelToTxt {
	private BufferedWriter out;
	
	public ExcelToTxt() throws IOException {
		out = new BufferedWriter(new FileWriter("out.txt"));
		readExcel("./data/서울시CCTV정보.xls");
		out.close();
	}
	
	public void readExcel(String file) {
		XSSFRow row;
		XSSFCell cell;
		
		try {
			XSSFWorkbook workbook = new XSSFWorkbook(file);
			
			XSSFSheet sheet = workbook.getSheetAt(0);
			int rows = sheet.getPhysicalNumberOfRows();
			int cells = sheet.getRow(0).getPhysicalNumberOfCells();
			
			for (int r = 0; r < rows; r++) {
                row = sheet.getRow(r); //  행 가져오기
                if (row != null) {
                    for (int c = 0; c < cells; c++) {
                        cell = row.getCell(c);
                        if (cell != null) {
                            String value = null;
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
                                value = "";
                                break;
                            case ERROR:
                                value = "" + cell.getErrorCellValue();
                                break;
                            default:
                            }
                           //System.out.print(value + ",");  //Test Print
                           writeFile(value);
                        } else {
                            
                        }
                    } 
                    writeFile("\n");
                    //System.out.println(); //Test Print
                }
            } 
			
			
		} catch (Exception e) {
			e.printStackTrace();
		}
		
	}
	
	public void writeFile(String str) throws IOException {
		str = ","+str;
		out.write(str);
	}
	
	public static void main(String[] args) {
		try {
			new ExcelToTxt();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
