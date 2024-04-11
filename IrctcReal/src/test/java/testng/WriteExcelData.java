package testng;

import org.testng.annotations.Test;

import utility.ExcelFileUtil;

public class WriteExcelData {
  @Test
  public void varifyexcelwritedata() throws Exception {
	  ExcelFileUtil util =new ExcelFileUtil();
	  util.ExcelFileUtilwrite("userid",0 , 4, 5, "bye bye");
	  int rows = util.getRowCount();
	  int cols = util.getColumnsCount(0);
	  
	 System.out.println(rows);
	 System.out.println(cols);
  }
  
}
