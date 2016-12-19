import java.io.File;
import java.io.FileInputStream;
import java.io.FilenameFilter;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.Timestamp;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import lombok.extern.slf4j.Slf4j;

@Slf4j
public class BaseDataInsert	{
	/**
	 * @param args
	 * @throws Exception
	 */
	public static void main(String[] args) throws Exception {
		String DB_URL = SystemValue.DB_URL;
		String DB_USER = SystemValue.DB_USER;
		String DB_PASSWORD = SystemValue.DB_PASSWORD;
		String EXCEL_DATA_FORDER = SystemValue.EXCEL_DATA_FORDER;
		if(args != null && args.length == 4)	{
			DB_URL = "jdbc:oracle:thin:@" + args[0] + ":1521:ORCL";
			DB_USER = args[1];
			DB_PASSWORD = args[2];
			EXCEL_DATA_FORDER = args[3];
		}

		try {
			// 드라이버를 로딩한다.
			Class.forName("oracle.jdbc.driver.OracleDriver");
		} catch (ClassNotFoundException e ) {
			e.printStackTrace();
		}

    File processDir = new File(EXCEL_DATA_FORDER);
		File[] fileList = processDir.listFiles(new FilenameFilter(){
			@Override
			public boolean accept(File dir, String name) {
				if(name.lastIndexOf('.')>0)	{
	                  int lastIndex = name.lastIndexOf('.');
	                  String str = name.substring(lastIndex);
	                  if(".xlsx".equals(str) && name.indexOf("TB_") > -1)	{
	                     return true;
	                  }
	               }
				return false;
			}
		});

		Connection conn = null;

		try{
			conn = DriverManager.getConnection(DB_URL, DB_USER, DB_PASSWORD);
			conn.setAutoCommit(false);
			PreparedStatement pstmt = null;

    		for(int i = 0 ; i < fileList.length ; i++)	{

				FileInputStream inputStream = new FileInputStream(fileList[i]);
				String tableName = fileList[i].getName();
				tableName = tableName.substring(0, tableName.lastIndexOf("."));
				if(tableName.indexOf("TB_") > 0)	{
					tableName = tableName.substring(tableName.indexOf("TB_"), tableName.length());
				}
				log.info("tableName = {}", tableName);

				pstmt = conn.prepareStatement("SELECT COLUMN_NAME FROM USER_TAB_COLUMNS WHERE TABLE_NAME = '" + tableName + "' AND DATA_TYPE = 'DATE'");
				ResultSet rs = pstmt.executeQuery();
				ArrayList<String> dateColList = new ArrayList<>();
				while(rs.next()) {
				    String name = rs.getString(1);
				    dateColList.add(name);
				}
				System.out.println(dateColList.size());
				dateColList.stream().forEach(System.out::println);


				ArrayList<String> column = new ArrayList<String>();
				StringBuffer columnName = new StringBuffer();
				StringBuffer prepareColumnData = new StringBuffer();
				int j = 0;

		        Workbook workbook = new XSSFWorkbook(inputStream);
		        Sheet firstSheet = workbook.getSheetAt(0);
		        Iterator<Row> iterator = firstSheet.iterator();

		        while (iterator.hasNext()) {
		            Row nextRow = iterator.next();
		            ArrayList<String> columnData = new ArrayList<String>();
		            LocalDateTime currentTime = LocalDateTime.now();
		            Timestamp t = Timestamp.valueOf(currentTime);
		            for(int k = 0; k < nextRow.getLastCellNum() ; k++)	{
		            	Cell cell = nextRow.getCell(k);
		            	if(j == 0)	{
		                	column.add(k, getCellValue(cell));
		                	columnName.append(", ").append(getCellValue(cell));
		                	prepareColumnData.append(", ?");
		                }else	{
		                	columnData.add(k, getCellValue(cell));
		                }
		            }

		            if(j == 0)	{
		            	pstmt = conn.prepareStatement("INSERT INTO " + tableName + " (" + columnName.substring(1) + " )" + " VALUES (" + prepareColumnData.substring(1) + " )");
		            }else	{
		            	for(int k = 0 ; k < column.size() ; k++)	{
		            		String c_name = column.get(k);
		            		String c_value = columnData.get(k);
		            		if(c_name != null)	{
		            			if(dateColList.size() > 0 && dateColList.stream().anyMatch(p -> p.equals(c_name)))	{
			            			if(c_name.indexOf("CREATE_DATE") > -1 || c_name.indexOf("MODIFY_DATE") > -1 || "CREATED".equals(c_name) || "UPDATED".equals(c_name))	{
				            			pstmt.setTimestamp(k + 1, t);
				            		}else if(c_name.indexOf("DATE") > -1 && c_value.length() >= 8)	{
				            			DateTimeFormatter format;
				            			if(c_value.length() == 8)	{
				            				format = DateTimeFormatter.ofPattern("yyyyMMdd");
				            			}else	{
				            				format = DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss");
				            			}
				            			LocalDate date = LocalDate.parse(c_value, format);
				            			pstmt.setDate(k + 1, java.sql.Date.valueOf(date));
				            		}
		            			}else 	{
			            			pstmt.setString(k + 1, c_value);
			            		}
		            		}
		            	}
		            	 // addBatch에 담기
		                pstmt.addBatch();
		                // 파라미터 Clear
		                pstmt.clearParameters() ;
		            }
		            j = 1;
		        }
		        // Batch 실행
		        pstmt.executeBatch();
				// Batch 초기화
				pstmt.clearBatch();
				conn.commit();

		        workbook.close();
		        inputStream.close();
    		}
		}catch (Exception e) {
			e.printStackTrace();
			if(conn != null)	{
				conn.rollback();
			}
		}finally	{
			if(conn != null)	{
				conn.commit();
				conn.close();
			}
		}
	}

	private static String getCellValue(Cell cell) {
		if(cell != null)	{
		    switch (cell.getCellType()) {
		    case Cell.CELL_TYPE_STRING:
		        return cell.getStringCellValue();

		    case Cell.CELL_TYPE_BOOLEAN:
		        return String.valueOf(cell.getBooleanCellValue());

		    case Cell.CELL_TYPE_NUMERIC:
		        return String.valueOf(cell.getNumericCellValue());
		    case Cell.CELL_TYPE_BLANK:
		        return "";
		    }
		}
	    return "";
	}
}
