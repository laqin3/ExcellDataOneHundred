package oneHundred.ExcelDataOneHundred;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteDataInExcel1 {

	public static void main(String[] args) {

		XSSFWorkbook workbook1 = new XSSFWorkbook();
		XSSFSheet sheet1 = workbook1.createSheet("first sheet");

		Map<String, Object[]> data = new HashMap<String, Object[]>();
		data.put("0", new Object[] { "name", "age", "salary" });
		data.put("1", new Object[] { "john", 45, 785123d });
		data.put("2", new Object[] { "jack", 28, 456213 });

		Set<String> keys = data.keySet();

		int rownum = 0;
		for (String key : keys) {
			XSSFRow row = sheet1.createRow(rownum++);

			Object[] values = data.get(key);
			int column = 0;
			for (Object value : values) {
				XSSFCell cell = row.createCell(column++);
				if (value instanceof Boolean) {
					cell.setCellValue((Boolean) value);
				}
				if(value instanceof Double) {
					cell.setCellValue((Double)value);
				}
				if(value instanceof String) {
					cell.setCellValue((String)value);
				}
				if(value instanceof Integer) {
					cell.setCellValue((Integer)value);
				}
			}
			
			

		}

		try {
			
			FileOutputStream out1=new FileOutputStream("C:\\Users\\laqin3\\Desktop\\dataDriven\\writtenExcel.xlsx");
			workbook1.write(out1);
			out1.close();
			System.out.println("Excel Written Successfully ..");
		
		
		
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

}