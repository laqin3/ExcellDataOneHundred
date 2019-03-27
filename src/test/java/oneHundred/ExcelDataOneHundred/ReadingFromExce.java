package oneHundred.ExcelDataOneHundred;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

public class ReadingFromExce {

	@Test(enabled = false)
	public void readingExcel() throws IOException {
		try {
			FileInputStream fis = new FileInputStream("C:\\Users\\laqin3\\Desktop\\dataDriven\\test.xlsx");

			// get workbook instance for XLSX file
			XSSFWorkbook workbook = new XSSFWorkbook(fis);

			// Get first sheet from the workbook
			XSSFSheet sheet = workbook.getSheetAt(0);

			// Iterate through each rows from first sheet
			Iterator<Row> rowiterator = sheet.iterator();
			while (rowiterator.hasNext()) {
				Row row = rowiterator.next();

				// For each row,iterate through each columns
				Iterator<Cell> cell = row.iterator();
				while (cell.hasNext()) {
					Cell ce = cell.next();

					switch (ce.getCellType()) {
					case BOOLEAN:
						System.out.println(ce.getBooleanCellValue() + "\t\t");// t--->type?
						break;
					case STRING:
						System.out.println(ce.getStringCellValue() + "\t\t");
						break;
					case NUMERIC:
						System.out.println(ce.getNumericCellValue() + "\t\t");
						break;
					}

				}
				System.out.println("");
			}
			fis.close();
			FileOutputStream out = new FileOutputStream("C:\\Users\\laqin3\\Desktop\\dataDriven\\test1.xlsx");// create
			workbook.write(out);
			out.close();

		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	@Test(enabled = false)
	public void creatNewExcelFile() {
		XSSFWorkbook wbook = new XSSFWorkbook();
		XSSFSheet wsheet = wbook.createSheet("Sample Sheet");
		XSSFRow wrow = wsheet.createRow(2);
		XSSFCell wcell = wrow.createCell(3);

		wcell.setCellValue("one hundred");
	}

	@Test
	public void writeDataInExcel() {
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet();

		// sheet.createRow(rownum);
		Map<String, Object[]> data = new HashMap<String, Object[]>();
		data.put("0", new Object[] { "Emp NO.", "Name", "Salary" });
		data.put("1", new Object[] { "1d", "John", 1500000d });
		data.put("2", new Object[] { "2d", "Sam", 800000d });
		data.put("3", new Object[] { "3d", "Dean", 18596456d });

		Set<String> keyset = data.keySet();// Returns a Set view of the keys contained in this map.The set is backed by
											// the map, setka?
		int rownum = 0;
		for (String key : keyset) {
			Row row = sheet.createRow(rownum++);
			Object[] objarr = data.get(key);
			int column = 0;
			for (Object obj : objarr) {
				Cell cell = row.createCell(column++);
				if (obj instanceof Date) {
					cell.setCellValue((Date) obj);
				}
				if (obj instanceof Boolean) {
					cell.setCellValue((Boolean) obj);
				}
				if (obj instanceof String) {
					cell.setCellValue((String) obj);
				}
				if (obj instanceof Double) {
					cell.setCellValue((Double) obj);
				}
			}
		
			try {
				FileOutputStream out1 = new FileOutputStream("C:\\Users\\laqin3\\Desktop\\dataDriven\\write1.xlsx");
				workbook.write(out1);
				out1.close();
				System.out.println(column+"Excel writtren Successfully..");
			} catch (FileNotFoundException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}

		}
	}

}
