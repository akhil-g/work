package excelread;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;

public class XLSXReader {
	private static Workbook wb1;
	private static Sheet sh1;
	private static FileInputStream fis1;
	private static Workbook wb2;
	private static Sheet sh2;
	private static FileInputStream fis2;	
	private static FileOutputStream fos;
	private static Row row;
	
	public static void main(String[] args)throws Exception {
		fis1 = new FileInputStream("./java_data.xlsx");
		wb1 = WorkbookFactory.create(fis1);
		sh1 = wb1.getSheet("Sheet1");
		fis2 = new FileInputStream("./output_data.xlsx");
		wb2 = WorkbookFactory.create(fis2);
		sh2 = wb2.getSheet("Sheet1");
		int noOfRows = sh1.getLastRowNum();
		ArrayList<String> al = new ArrayList<String>();
		
		//EXCEL READ OPERATIONS CODE
		for(int i=0; i<=noOfRows;i++) {
			for(int j=0;j<sh1.getRow(i).getPhysicalNumberOfCells();j++) {
				final DataFormatter df = new DataFormatter();
				final XSSFCell cell = (XSSFCell) sh1.getRow(i).getCell(j);
				String valueAsString = df.formatCellValue(cell);
				al.add(valueAsString);
			}
		}
		//EXCEL WRITE OPERATION CODE
		int k = 0;
		for(int i=0; i<=noOfRows;i++) {
			row = sh2.createRow(i);
			for(int j=0;j<sh1.getRow(i).getPhysicalNumberOfCells();j++) {
				String str = (String) al.get(k);
				row.createCell(j).setCellValue(str);
				k += 1;
			}
		}
		fos = new FileOutputStream("./output_data.xlsx");
		wb2.write(fos);
		fos.flush();
	}
}
