import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.*;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import com.mysql.jdbc.Connection;

public class TestHSSFWorkbook {
	public static void main(String[] args) throws Exception {
		TestHSSFWorkbook tm = new TestHSSFWorkbook();
		tm.jdbcex(true);
	}
	public void jdbcex(boolean isClose) throws InstantiationException, IllegalAccessException, 
				ClassNotFoundException, SQLException, IOException, InterruptedException {
			
		String xlsFile = "e:/poiBigData.xls";		//输出文件
		//内存中只创建100个对象，写临时文件，当超过100条，就将内存中不用的对象释放。
		HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream( xlsFile));
		
		Sheet sheet = null;		//工作表对象
		Row nRow = null;		//行对象
		Cell nCell = null;		//列对象

		//使用jdbc链接数据库
		Class.forName("com.mysql.jdbc.Driver").newInstance();  
		String url = "jdbc:mysql://localhost:3306/test?characterEncoding=UTF-8";
		String user = "root";
		String password = "root";
		long  startTime = System.currentTimeMillis();	//开始时间
		System.out.println("strat execute time: " + startTime);
		//获取数据库连接
		Connection conn = (Connection)DriverManager.getConnection(url, user,password);   
		Statement stmt = conn.createStatement(ResultSet.TYPE_SCROLL_SENSITIVE,ResultSet.CONCUR_UPDATABLE);   
		String sql = "select * from hpa_normal_tissue limit 250000";   //100万测试数据
		ResultSet rs = stmt.executeQuery(sql);  
		
		ResultSetMetaData rsmd = rs.getMetaData();
		
			
		int rowNo = 0;		//总行号
		int pageRowNo = 0;	//页行号,初始行号
			
		while(rs.next()) {
			//打印300000条后切换到下个工作表，可根据需要自行拓展，2百万，3百万...数据一样操作，只要不超过1048576就可以
			if(rowNo%30000==0){
				System.out.println("Current Sheet:" + rowNo/30000);
			
				sheet = workbook.createSheet();//建立新的sheet对象
				System.out.print("我的第"+(rowNo/30000)+"个工作簿");
				sheet = workbook.getSheetAt(rowNo/30000);		//动态指定当前的工作表
				pageRowNo = 0;		//每当新建了工作表就将当前工作表的行号重置为0
			}	
			rowNo++;
			nRow = sheet.createRow(pageRowNo++);	//新建行对象

			// 打印每行，每行有6列数据   rsmd.getColumnCount()==6 --- 列属性的个数
			for(int j=0;j<rsmd.getColumnCount();j++){
				nCell = nRow.createCell(j);
				nCell.setCellValue(rs.getString(j+1));
			}
				
			/*if(rowNo%10000==0){
				System.out.println("row no: " + rowNo);
			}*/
//			Thread.sleep(1);	//休息一下，防止对CPU占用，其实影响不大
		}
			
		long finishedTime = System.currentTimeMillis();	//处理完成时间
		System.out.println("finished execute  time: " + (finishedTime - startTime)/1000 + "m");
			
		FileOutputStream fOut = new FileOutputStream(xlsFile);
		workbook.write(fOut);
		fOut.flush();		//刷新缓冲区
		fOut.close();
			
		long stopTime = System.currentTimeMillis();		//写文件时间
		System.out.println("write xlsx file time: " + (stopTime - startTime)/1000 + "m");
			
		if(isClose){
			this.close(rs, stmt, conn);
		}
	}
		
	//执行关闭流的操作
	private void close(ResultSet rs, Statement stmt, Connection conn ) throws SQLException{
		rs.close();   
		stmt.close();   
		conn.close(); 
	}
}
