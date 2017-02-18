

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.Statement;



public class JDBCTest {
	public static void main(String[] args) throws Exception {
		System.out.println("Hi111");
		getDepartmenList();
		System.out.println("Hi22");
	}
	
	public static void getDepartmenList() throws Exception{
		//Driver class setup
		Class.forName("com.mysql.jdbc.Driver");  
		System.out.println("Hi");
		//Database connection establishment
		Connection con=DriverManager.getConnection("jdbc:mysql://localhost:3306/pmec","root","root");  
		
		//Connection object creation
		Statement stmt=con.createStatement();
	    String query="SELECT * FROM department";
	    ResultSet rs=stmt.executeQuery(query);
	    while(rs.next()) {
	        System.out.print(rs.getString("DepartmentId") +" ");
	        System.out.print(rs.getString("DepartmentName") +" ");
	        System.out.println("");
	    } 
	    
	    //Closing connection
	    rs.close();
	    stmt.close();
	    con.close();
	}
}	        


