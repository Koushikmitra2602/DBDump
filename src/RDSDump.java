import java.io.FileOutputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.Statement;
import java.util.List;
import java.util.Map;
import java.util.Properties;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class RDSDump {

	@SuppressWarnings("unchecked")
	public static void main(String[] args) {

		//get the connection
		Connection connection;
		try {
			Map<String, Object> inputList = ExcelReader.getTableList();
			Map<String, String> propMap = (Map<String, String>) inputList.get("PROP");
			connection = getDBConnectionUsingIam(propMap);
			//verify the connection is successful
			Statement stmt= connection.createStatement();
			Workbook workbook = new XSSFWorkbook(); // new HSSFWorkbook() for generating `.xls` file
			
			for(String tableName : (List<String>) inputList.get("TABLES")) {
				ResultSet rs=stmt.executeQuery("SELECT * FROM "+ tableName);
				ExcelWriter.writeExcel(rs, workbook);
			}

			// Write the output to a file
			FileOutputStream fileOut = new FileOutputStream(propMap.get("DatabaseName")+".xlsx");
			workbook.write(fileOut);
			fileOut.close();

			// Closing the workbook
			workbook.close();

			//close the connection
			stmt.close();
			connection.close();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

	/**
	 * This method returns a connection to the db instance authenticated using IAM Database Authentication
	 * @return
	 * @throws Exception
	 */
	private static Connection getDBConnectionUsingIam(Map<String, String> propMap) throws Exception {
		return DriverManager.getConnection("jdbc:"+propMap.get("DatabaseType")+"://" + propMap.get("Host") +":"+propMap.get("Port")+"/"+propMap.get("DatabaseName"), setMySqlConnectionProperties());
	}

	/**
	 * This method sets the mysql connection properties which includes the IAM Database Authentication token
	 * as the password. It also specifies that SSL verification is required.
	 * @return
	 */
	private static Properties setMySqlConnectionProperties() {
		Properties mysqlConnectionProperties = new Properties();
		mysqlConnectionProperties.setProperty("user","root");
		mysqlConnectionProperties.setProperty("password","20170908");
		return mysqlConnectionProperties;
	}

}

