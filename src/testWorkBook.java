import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFCellUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

//读取文件 批量导出文件
public class testWorkBook {

	public static void main(String[] args) throws InstantiationException, IllegalAccessException, 
	ClassNotFoundException, SQLException, IOException, InterruptedException{
		long beginTime = System.currentTimeMillis();
        xlsx_reader();
        long endTime = System.currentTimeMillis();
        System.out.println("SXSSFWorkbook运行时间："+ (endTime - beginTime));
        
        long beginTime1 = System.currentTimeMillis();
        Excel2007AboveOperateOld();
        long endTime1 = System.currentTimeMillis();
        System.out.println("XSSFWorkbook运行时间:"+ (endTime1 - beginTime1));
	}
	//SXSSFWorkbook批量导出时间
	public static void xlsx_reader() throws IOException {       
		//支持读
        XSSFWorkbook workbook1 = new XSSFWorkbook(new FileInputStream(new File("E:\\work\\test.xlsx")));
        SXSSFWorkbook sxssfWorkbook = new SXSSFWorkbook(workbook1, 10000);//不支持读操作
        Sheet first = sxssfWorkbook.getSheetAt(0);
        int num = first.getLastRowNum();
        System.out.println(num);
        for (int i = 0; i < 100000; i++) {
            Row row = first.createRow(i);
            for (int j = 0; j < 15; j++) {
                if(i == 0) {
                    // 首行
                    row.createCell(j).setCellValue("column" + j);
                } else {
                    // 数据
                    if (j == 0) {
                        CellUtil.createCell(row, j, String.valueOf(i));
                    } else
                        CellUtil.createCell(row, j, String.valueOf(Math.random()));
                }
            }
        }
        FileOutputStream out = new FileOutputStream("E:\\work\\test1.xlsx");
        sxssfWorkbook.write(out);
        out.close();
        
	}
	//XSSFWorkbook批量导出时间
    public static void Excel2007AboveOperateOld() throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(new File("E:\\work\\test.xlsx")));
        // 获取第一个表单
        Sheet first = workbook.getSheetAt(0);
        for (int i = 0; i < 100000; i++) {
            Row row = first.createRow(i);
            for (int j = 0; j < 15; j++) {
                if(i == 0) {
                    // 首行
                    row.createCell(j).setCellValue("column" + j);
                } else {
                    // 数据
                    if (j == 0) {
                        CellUtil.createCell(row, j, String.valueOf(i));
                    } else
                        CellUtil.createCell(row, j, String.valueOf(Math.random()));
                }
            }
        }
        // 写入文件
        FileOutputStream out = new FileOutputStream("E:\\work\\test2.xlsx");
        workbook.write(out);
        out.close();
    }
    //HSSFWorkbook批量导出时间
    public static void Excel2003Operate(String filePath) throws Exception {
        HSSFWorkbook hssfWorkbook = new HSSFWorkbook(new FileInputStream(new File("E:\\work\\test.xlsx")));
        HSSFSheet sheet = hssfWorkbook.getSheetAt(0);
        for (int i = 0; i < 10000; i++) {
            HSSFRow hssfRow = sheet.createRow(i);
            for (int j = 0; j < 10; j++) {
                HSSFCellUtil.createCell(hssfRow, j, String.valueOf(Math.random()));
            }
        }
        FileOutputStream out = new FileOutputStream("E:\\work\\test2.xlsx");
        hssfWorkbook.write(out);
        out.close();
    }


}
