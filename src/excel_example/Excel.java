package excel_example;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel {
	
	private static final String filename = "data\\workbook.xlsx";
	private static Workbook wb;
	private static Sheet sheet1;
	static Object[][] bookData = {
            {"Head First Java", "Kathy Serria", 79},
            {"Effective Java", "Joshua Bloch", 36},
            {"Clean Code", "Robert martin", 42},
            {"Thinking in Java", "Bruce Eckel", 35},
    	};
	static int rowCount = 0;
		
	
		public static void main(String[] args) {
			wb = new XSSFWorkbook();
			sheet1 = wb.createSheet("new sheet");
			
			for (Object[] aBook : bookData) {
	            Row row = sheet1.createRow(++rowCount);
	             
	            int columnCount = 0;
	             
	            for (Object field : aBook) {
	                Cell cell = row.createCell(++columnCount);
	                if (field instanceof String) {
	                    cell.setCellValue((String) field);
	                } else if (field instanceof Integer) {
	                    cell.setCellValue((Integer) field);
	                }
	            }
	             
	        }
			
			try {
				FileOutputStream fileOut = new FileOutputStream(filename);
		        wb.write(fileOut);
		        fileOut.close();
			}
			catch(IOException ex) {
				System.err.println("An IOException was caught!");
				ex.printStackTrace();
			}
			
			}
		}

	



