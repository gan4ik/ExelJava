package Main;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;

public class JavaExcelApp {

	public static void main(String[] args) throws IOException {
		
		Workbook wb = new HSSFWorkbook();
		Sheet sheet0 = wb.createSheet("��������");
		Row row = sheet0.createRow(3);
		Cell cell = row.createCell(4);
		cell.setCellValue("O'Reilly");
		
		Sheet sheet1 = wb.createSheet("�����");
		Row row1 = sheet1.createRow(0);
		Cell Cell1 = row1.createCell(0);
		Cell1.setCellValue("����� � ���");
		
		Row row2 = sheet1.createRow(1);
		Cell cell2 = row2.createCell(3);
		cell2.setCellValue("������� ������");
		
		Sheet sheet2 = wb.createSheet("������");
		Sheet sheet3 = wb.createSheet(WorkbookUtil.createSafeSheetName("sfh�!#$"));
		
		FileOutputStream fos = new FileOutputStream("C:/Users/GannIV.GU.003/workspace/JavaExel/my.xls");
		
		wb.write(fos);
		fos.close();
		
	}

}
