import java.sql.*;

import com.sun.corba.se.spi.orbutil.fsm.Guard.Result;

public class DBaccess {

	static String table1 = "Shift Handed Over";
	static Float inflowOutflowRatio,inflow,outflow,memberAverageResolvedBatch,memberAverageResolvedNonBatch;
	public static void main(String[] args) {
		// TODO Auto-generated method stub
try{
	Class.forName("net.ucanaccess.jdbc.UcanaccessDriver");
	Connection conn=DriverManager.getConnection(
	        "jdbc:ucanaccess://C:/Users/vijsu/Desktop/Shift Summary Report/Shift Summary Report.accdb");
	Statement s = conn.createStatement();
	ResultSet countFromPrevShift = s.executeQuery("SELECT COUNT(Number) FROM [Total Count from prev shift]");
	ResultSet countHandedOverToNextShift = s.executeQuery("SELECT COUNT(Number) FROM [Shift Handed Over]");
	
	while (countFromPrevShift.next()) {
	    System.out.println("Count From Previous Shift: "+countFromPrevShift.getFloat(1));
	    inflow = countFromPrevShift.getFloat(1);
	}
	while (countHandedOverToNextShift.next()) {
	    System.out.println("Count From Handedover to Next Shift: "+countHandedOverToNextShift.getFloat(1));
	    outflow = countHandedOverToNextShift.getFloat(1);
	}
	inflowOutflowRatio = inflow/outflow;
	System.out.println("Inflow / Outflow Ratio : "+inflowOutflowRatio);
	//caluclations(countFromPrevShift,countHandedOverToNextShift);
    
}
catch(Exception ex){
	ex.printStackTrace();
}
	}
	private static void caluclations(ResultSet countFromPrevShift, ResultSet countHandedOverToNextShift) {
		// TODO Auto-generated method stub
		
	}

}
