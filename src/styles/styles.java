package styles;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class styles {

	public static void main(String[] args) throws IOException {

		Workbook wb = new HSSFWorkbook();
		Sheet sheet0 = wb.createSheet("Стили");
		Row row = sheet0.createRow(0);		
		Cell cell0 = row.createCell(0);
		cell0.setCellValue("Hi");
		
		CellStyle style = wb.createCellStyle();
		style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
		style.setFillBackgroundColor(IndexedColors.GREEN.getIndex());
		
		
        FileOutputStream fos = new FileOutputStream("C:/Users/GannIV.GU.003/workspace/JavaExel/styles.xls");
		
		wb.write(fos);
		fos.close();
		wb.close();

	}

}
