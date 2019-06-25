import java.io.File;
import java.io.FileInputStream;
import java.util.Date;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.format.CellDateFormatter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		// TODO Auto-generated method stub

		try {
			FileInputStream file = new FileInputStream(new File("d:\\test.xlsx"));

			// Create Workbook instance holding reference to .xlsx file
			XSSFWorkbook workbook = new XSSFWorkbook(file);

			// Get first/desired sheet from the workbook
			XSSFSheet sheet = workbook.getSheetAt(0);

			// Iterate through each rows one by one
			Iterator<Row> rowIterator = sheet.iterator();
			StringBuilder sqlBuilder = new StringBuilder();
			String readLine = "";
			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();
				// For each row, iterate through all the columns
				Iterator<Cell> cellIterator = row.cellIterator();
				readLine = "";
				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();
					// Check the cell type and format accordingly
					switch (cell.getCellType()) {
					case Cell.CELL_TYPE_NUMERIC:
						if (HSSFDateUtil.isCellDateFormatted(cell)) {
							double dv = cell.getNumericCellValue();
							Date datap = HSSFDateUtil.getJavaDate(dv);
							String df = cell.getCellStyle().getDataFormatString();
							String reportDatap = new CellDateFormatter(df).format(datap);
							readLine += reportDatap + ",";
						} else {
							readLine += cell.getNumericCellValue() + ",";
						}
						break;
					case Cell.CELL_TYPE_STRING:
						readLine += cell.getStringCellValue() + ",";
						break;
					}
				}
				sqlBuilder.append(readLine.substring(0, readLine.length() - 1));
				sqlBuilder.append(System.getProperty("line.separator"));
				// System.out.println("");
			}
			System.out.println(sqlBuilder.toString());
			file.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void InsertRowInDB(String name, String empid, String add, String mobile) {

	}

}
