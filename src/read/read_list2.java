package read;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;

public class read_list2 {
	
	public static SimpleDateFormat sdf = new SimpleDateFormat("yyyy.MM.dd");

	public static void main(String[] args) throws IOException {
		
		FileInputStream fis = new FileInputStream("C:/scripts/�� ����������.xls");
		Workbook wb = new HSSFWorkbook(fis);
		/*
		String result = wb.getSheetAt(0).getRow(0).getCell(0).getStringCellValue();
		System.out.println(result);
		
		System.out.println(getCellText(wb.getSheetAt(0).getRow(0).getCell(1)));
		System.out.println(getCellText(wb.getSheetAt(0).getRow(0).getCell(2)));
		System.out.println(getCellText(wb.getSheetAt(0).getRow(0).getCell(3)));
		*/
		
		for(Row row : wb.getSheetAt(0)){
			for (Cell cell : row){
				
				CellReference cellRef = new CellReference(row.getRowNum(), cell.getColumnIndex());
	            System.out.print(cellRef.formatAsString());
	            //System.out.print(" - ");
				
	            System.out.println(getCellText(cell));
			
			}
		}
		
		fis.close(); 
	}
	
	public static String getCellText(Cell cell) {
		
		String result="";
		
		switch (cell.getCellType ()) {
		
        case Cell.CELL_TYPE_STRING:
            result = cell.getRichStringCellValue().getString();
            break;
        case Cell.CELL_TYPE_NUMERIC:
            if (DateUtil.isCellDateFormatted(cell)) {
                result = sdf.format(cell.getDateCellValue());
            } else {
                result = Double.toString(cell.getNumericCellValue());
            }
            break;
        case Cell.CELL_TYPE_BOOLEAN:
            result = Boolean.toString(cell.getBooleanCellValue());
            break;
        case Cell.CELL_TYPE_FORMULA:
            result = cell.getCellFormula().toString();
            break;
        /*case Cell.CELL_TYPE_BLANK:
            System.out.println();
            break;*/
        default:
            //System.out.println();
    }
		return result;
		
	}

}
