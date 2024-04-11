package utility;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;




public class ExcelFileUtil {

	public FileInputStream fis ;
	public FileOutputStream fos;
	public File file;
	public File file1;
	public XSSFWorkbook wb;
	public  Sheet sheet ;
	
	public  String ExcelFileUtil(String fileName , String sheetName ,int rNum,int cNum) throws Exception {
		//file =new File(System.getProperty("user.dir")+"/TestData/"+"userid"+".xlsx");
		String data ="";
		fis =new FileInputStream("C:\\Users\\niraj\\eclipse-workspace\\IrctcReal\\TestData\\userid.xlsx");
		wb =new XSSFWorkbook(fis);
		sheet= wb.getSheet(sheetName);
		Row r =sheet.getRow(rNum);
		Cell c =r.createCell(cNum);
		return data;
		
	}
	
	public void ExcelFileUtilwrite(String fileName , int sNum , int rNum , int cNum, String Data) throws Exception {
		//file =new File(System.getProperty("./TestData"+"userid"+".xlsx");
		fis =new FileInputStream("C:\\Users\\niraj\\eclipse-workspace\\IrctcReal\\TestData\\userid.xlsx");
		wb =new XSSFWorkbook(fis);
		sheet= wb.getSheetAt(sNum);
		Row r =sheet.getRow(rNum);
		Cell c =r.createCell(cNum);
		c.setCellValue(Data);
		file1 =new File(System.getProperty("user.dir")+"//TestDataOutput//" + "result1" +".xlsx");
		fos =new FileOutputStream(file1);
		wb.write(fos);
	}
	public int getRowCount() {
		return sheet.getPhysicalNumberOfRows();
	}
	public int getColumnsCount(int rowNumber) {
		return sheet.getRow(rowNumber).getPhysicalNumberOfCells();
	}
	public  String  getCellValue(int rowNumber , int columnNumber) {
		return sheet.getRow(rowNumber).getCell(columnNumber).getStringCellValue();
	}
	
	
}

