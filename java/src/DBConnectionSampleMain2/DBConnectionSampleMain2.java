import java.sql.ResultSet;
import java.sql.Statement;
import java.sql.Connection;
import java.sql.DriverManager;

public class DBConnectionSampleMain2 {
	public static void main(String[] args) {
		String framework = "embedded";
		String driver = "org.apache.derby.jdbc.EmbeddedDriver";
		String protocol = "jdbc:derby://localhost:1527/";
		try{
			//Class.forName(driver).newInstance();
			System.out.println("Derby DB Driver is loaded");
			Connection conn = null;
			conn = DriverManager.getConnection(protocol + "C:/Users/oyaji/Documents/mydata/db/derby/DerbyTestDB");
			Statement stmt = conn.createStatement();
			ResultSet rs = stmt.executeQuery("select * from member");
			while(rs.next()) {
				System.out.println(
					rs.getString(1) +
					"," + 
					rs.getString(2)
					);
			}
			rs.close();
			stmt.close();
			conn.close();
			System.out.println("fin");
		} catch(Exception e) {
			e.printStackTrace();
		}
	}
}
