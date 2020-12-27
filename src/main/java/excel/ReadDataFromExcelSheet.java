package excel;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadDataFromExcelSheet {
	public String path;
	public FileInputStream fis = null;
	public FileOutputStream fileOut = null;
	private XSSFWorkbook workbook = null;
	private XSSFSheet sheet = null;
	private XSSFRow row = null;
	private XSSFCell cell = null;

	ReadDataFromExcelSheet() {
	}

	public ReadDataFromExcelSheet(String path) {

		this.path = path;
		try {
			fis = new FileInputStream(path);
			workbook = new XSSFWorkbook(fis);
			sheet = workbook.getSheetAt(0);

			fis.close();
		} catch (Exception e) {

			e.printStackTrace();
		}

	}

	public List<Map<String, String>> getExcelData(String excellocation, String sheetName) {
		try {
			List<Map<String,String>> allData=new ArrayList<>();
			FileInputStream file = new FileInputStream(new File(excellocation));

			// Create Workbook instance holding reference to .xlsx file
			XSSFWorkbook workbook = new XSSFWorkbook(file);
//			workbook.setMissingCellPolicy(MissingCellPolicy.RETURN_NULL_AND_BLANK);

			// Get first/desired sheet from the workbook
			XSSFSheet sheet = workbook.getSheet(sheetName);
			// count number of active rows
			int totalRow = sheet.getLastRowNum();
			System.out.println("totalRow - " + (totalRow + 1));
			// count number of active columns in row
			int totalColumn = sheet.getRow(0).getLastCellNum();
			System.out.println("totalColumn - " + totalColumn);
			// Create array of rows and column
			// Iterate through each rows one by one
			Iterator<Row> rowIterator = sheet.iterator();
			int i = 1;
			while (rowIterator.hasNext()) {

				Row row = rowIterator.next();
				// For each row, iterate through all the columns
				Iterator<Cell> cellIterator = row.cellIterator();
				int j = 0;
				System.out.println("Row counter - "+i);
				Map<String,String> data=new LinkedHashMap<>();
				while (cellIterator.hasNext()) {
					if(i==1)
						break;

					Cell cell = cellIterator.next();
					
					switch (cell.getCellType()) {
					
					case STRING:
						System.out.println(String.valueOf(cell.getStringCellValue()));
						data.put(String.valueOf(sheet.getRow(0).getCell(j++).getStringCellValue().trim()),
								String.valueOf(cell.getStringCellValue()));
						break;
					case NUMERIC:
						System.out.println(String.valueOf(cell.getNumericCellValue()));
						data.put(String.valueOf(sheet.getRow(0).getCell(j++).getStringCellValue().trim()),
								String.valueOf(cell.getNumericCellValue()));
						break;
					case BOOLEAN:
						System.out.println(String.valueOf(cell.getBooleanCellValue()));
						data.put(String.valueOf(sheet.getRow(0).getCell(j++).getStringCellValue().trim()),
								String.valueOf(cell.getStringCellValue()));
						break;
					case BLANK:
						data.put(String.valueOf(sheet.getRow(0).getCell(j++).getStringCellValue().trim()),
								String.valueOf(cell.getStringCellValue()));
						break;
					case FORMULA:
						System.out.println(String.valueOf(cell.getStringCellValue()));
						data.put(String.valueOf(sheet.getRow(0).getCell(j++).getStringCellValue().trim()),
								String.valueOf(cell.getStringCellValue()));
						break;
					default:
						break;
				
				}

				}
				if(data.size()>0)
				{
					allData.add(data);
				}

				System.out.println("");
				i++;
			}
			file.close();
			workbook.close();
			
			return allData;
		} catch (Exception e) {
			e.printStackTrace();
		}
		return null;
	}

	// returns the row count in a sheet
	public int getRowCount(String sheetName) {
		int index = workbook.getSheetIndex(sheetName);
		if (index == -1)
			return 0;
		else {
			sheet = workbook.getSheetAt(index);
			/**
			 * getLastRowNum() always returns one less count as index starts from 0 so added
			 * 1
			 */

			int number = sheet.getLastRowNum() + 1;
			return number;
		}

	}

	// returns number of columns in a sheet
	public int getColumnCount(String sheetName) {
		// check if sheet exists
		if (!isSheetExist(sheetName))
			return -1;

		sheet = workbook.getSheet(sheetName);
		row = sheet.getRow(2);
		if (row == null)
			return -1;

		return row.getLastCellNum();

	}

	// find whether sheets exists
	public boolean isSheetExist(String sheetName) {
		int index = workbook.getSheetIndex(sheetName);
		if (index == -1) {
			index = workbook.getSheetIndex(sheetName.toUpperCase());
			if (index == -1)
				return false;
			else
				return true;
		} else
			return true;
	}
	public static void main(String[] args) {
		String excellocation = System.getProperty("user.dir")+File.separator+"shreyansh.xlsx";
		String sheetName = "Sheet1";
		
		ReadDataFromExcelSheet excelSheet = new ReadDataFromExcelSheet(excellocation);
		System.out.println(excelSheet.getExcelData(excellocation, sheetName));
		int rowCount = excelSheet.getRowCount(sheetName);
		int columnCount = excelSheet.getColumnCount(sheetName);
		
	}
	
}