import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class TestXSSFWorkbook {

	public static void main(String[] args) throws InstantiationException, IllegalAccessException, 
	ClassNotFoundException, SQLException, IOException, InterruptedException{
        xlsx_reader();
	}
	//导入文件操作
	public static void xlsx_reader() throws IOException {  
        //读取xlsx文件  
        XSSFWorkbook xssfWorkbook = null;  
        //寻找目录读取文件  
        File excelFile = new File("E:\\work\\test.xlsx");
        InputStream is = new FileInputStream(excelFile);  
        xssfWorkbook = new XSSFWorkbook(is);  
        XSSFSheet sheet = xssfWorkbook.getSheetAt(0);// 取出第一个工作表，索引是0  XSSFSheet是为来读取数据使用
        System.out.println(sheet.getSheetName());
        // 开始循环遍历行，表头不处理，从1开始
        for (int i = 0; i <= sheet.getLastRowNum(); i++) {
        	XSSFRow row = sheet.getRow(i);// 获取行对象
	        if (row == null) {// 如果为空，不处理
	          continue;
	        }
         // 循环遍历单元格
         for (int j = 0; j < row.getLastCellNum(); j++) {
	          XSSFCell cell = row.getCell(j);// 获取单元格对象
	          String cellStr = "";
	          
	          if (cell == null) {// 单元格为空设置cellStr为空串
	        	  cellStr = "";
	          } else if (cell.getCellType() == XSSFCell.CELL_TYPE_BOOLEAN) {// 对布尔值的处理
	        	  cellStr = String.valueOf(cell.getBooleanCellValue());
	          } else if (cell.getCellType() == XSSFCell.CELL_TYPE_NUMERIC) {// 对数字值的处理
	        	  cellStr = cell.getNumericCellValue() + "";
	          } else {// 其余按照字符串处理
	        	  cellStr = cell.getStringCellValue();
	          }
	          System.out.print(cellStr + "  ");
         	}
         	System.out.println();
        }
	}
}
