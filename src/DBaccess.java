import java.sql.*;
import java.io.*;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableColumn;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableColumns;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableStyleInfo;
import com.sun.corba.se.spi.orbutil.fsm.Guard.Result;

public class DBaccess {

	static String table1 = "Shift Handed Over";
	static Float inflowOutflowRatio, inflow, outflow, memberAverageResolvedBatch, memberAverageResolvedNonBatch,
			wipPercentage;
	static int toBeChecked, withAdvTeam, withOtherTeams, withUsers, incidentsClosedByEach, tasksClosedByEach;
	static ResultSet countFromPrevShift, countHandedOverToNextShift, toBeCheckedSet, withAdvTeamSet, withUsersSet,
			withOtherTeamsSet, tasksClosedByEachSet, batchTasksClosedByEachSet, incidentsClosedByEachSet,
			batchIncidentsClosedByEachSet;

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		try {
			Class.forName("net.ucanaccess.jdbc.UcanaccessDriver");
			Connection conn = DriverManager.getConnection(
					"jdbc:ucanaccess://C:/Users/vijsu/Desktop/Shift Summary Report/Shift Summary Report.accdb");
			Statement s = conn.createStatement();
			countFromPrevShift = s.executeQuery("SELECT COUNT(Number) FROM [Total Count from prev shift]");
			countHandedOverToNextShift = s.executeQuery("SELECT COUNT(Number) FROM [Shift Handed Over]");
			toBeCheckedSet = s.executeQuery("SELECT [Impacted Locations],Count(Number) FROM [Shift Handed Over] "
					+ "WHERE [Next Action] LIKE '%to be checked%' GROUP BY [Impacted Locations]");
			withAdvTeamSet = s.executeQuery("SELECT [Impacted Locations],Count(Number)FROM [Shift Handed Over] "
					+ "WHERE ([Next Action] LIKE '%Adv%' AND ([Next Action] LIKE '%OM%' OR [Next Action] LIKE '%DL%' "
					+ "OR [Next Action] LIKE '%OTC%'))GROUP BY [Impacted Locations]");
			withUsersSet = s.executeQuery("SELECT [Impacted Locations],Count(Number)FROM [Shift Handed Over] "
					+ "WHERE ([Next Action] LIKE '%user%')GROUP BY [Impacted Locations]");
			withOtherTeamsSet = s.executeQuery("SELECT [Impacted Locations],Count(Number)FROM [Shift Handed Over] "
					+ "WHERE (([Next Action] NOT LIKE ('%user%')) AND "
					+ "([Next Action] NOT LIKE ('%om%')) AND ([Next Action] NOT LIKE ('%to be checked%')) AND "
					+ "([Next Action] NOT LIKE ('%om%')) AND ([Next Action] NOT LIKE ('%DL%')) AND "
					+ "([Next Action] NOT LIKE ('%otc%')) AND ([Next Action] NOT LIKE ('%closed%')))GROUP BY [Impacted Locations]");
			tasksClosedByEachSet = s.executeQuery(
					"SELECT [Assigned To],COUNT(Number) FROM [Tasks Closed count from SNOW] GROUP BY [Assigned To]");
			incidentsClosedByEachSet = s.executeQuery("SELECT [Assigned To],COUNT (Number), "
					+ "COUNT(Number) FILTER(WHERE (([Short Description] NOT LIKE ('%autosys%')) AND ([Short Description] NOT LIKE ('%rebalanc%')) AND ([Short Description] NOT LIKE ('%monitor%')))) FROM [Incidents closed count from SNOW] GROUP BY [Assigned To]");

			while (countFromPrevShift.next()) {
				System.out.println("Count From Previous Shift: " + countFromPrevShift.getFloat(1));
				inflow = countFromPrevShift.getFloat(1);
			}
			while (countHandedOverToNextShift.next()) {
				System.out.println("Count From Handedover to Next Shift: " + countHandedOverToNextShift.getFloat(1));
				outflow = countHandedOverToNextShift.getFloat(1);
			}
			System.out.println("To Be Checked \n Location \t\t Count ");
			while (toBeCheckedSet.next()) {

				System.out.println(toBeCheckedSet.getString(1) + "\t\t" + toBeCheckedSet.getString(2));
				wipPercentage = toBeCheckedSet.getFloat(2);
			}
			System.out.println("With Adv Teams \n Location \t\t Count ");
			while (withAdvTeamSet.next()) {
				System.out.println(withAdvTeamSet.getString(1) + "\t\t" + withAdvTeamSet.getInt(2));
				withAdvTeam = withAdvTeamSet.getInt(2);
			}
			System.out.println("With Users \n Location \t\t Count ");
			while (withUsersSet.next()) {
				System.out.println(withUsersSet.getString(1) + "\t\t" + withUsersSet.getInt(2));
				withAdvTeam = withUsersSet.getInt(2);
			}
			System.out.println("With Other Teams \n Location \t\t Count ");
			while (withOtherTeamsSet.next()) {
				System.out.println(withOtherTeamsSet.getString(1) + "\t\t" + withOtherTeamsSet.getInt(2));
				withOtherTeams = withOtherTeamsSet.getInt(2);
			}
			System.out.println("Tasks closed by each \n Closed By \t\t Count ");
			while (tasksClosedByEachSet.next()) {
				System.out.println(tasksClosedByEachSet.getString(1) + "\t\t" + tasksClosedByEachSet.getInt(2));
				tasksClosedByEach = tasksClosedByEachSet.getInt(2);
			}

			/*
			 * System.out.println(
			 * "Incidents closed by each \n Closed By \t\t Count ");
			 * while(incidentsClosedByEachSet.next()){ System.out.println(
			 * "Size of result"+incidentsClosedByEachSet.getFetchSize());
			 * System.out.println(incidentsClosedByEachSet.getString(1)+"\t\t"+
			 * incidentsClosedByEachSet.getInt(2)); incidentsClosedByEach =
			 * incidentsClosedByEachSet.getInt(2); }
			 */

			inflowOutflowRatio = inflow / outflow;
			wipPercentage = (wipPercentage / outflow) * 100;
			System.out.println("Inflow / Outflow Ratio : \t" + inflowOutflowRatio);
			System.out.println("WIP % (Ideal Value <= 20): \t" + wipPercentage);
			// ----------------------------------------Creating Excel
			// Report-------------------------------------------//
			createExcelReport();
		} catch (Exception ex) {
			ex.printStackTrace();
		}
	}

	private static void createExcelReport() throws IOException, SQLException {
		// TODO Auto-generated method stub
		String filepath = "C:\\Users\\vijsu\\Desktop\\Shift Summary Report\\Shift Summary Report_v4.xlsx";
		try {
			FileInputStream file = new FileInputStream(filepath);
			XSSFWorkbook workbook = new XSSFWorkbook(file);
			XSSFSheet sheet = workbook.getSheetAt(0);
			// ------Update the Carry Forward From Previous Shift------//
			XSSFCell carryForwardFromPreviousShift = sheet.getRow(2).getCell(2);
			carryForwardFromPreviousShift.setCellValue(12);
			// ---Update Incidents closed by each member----//
			System.out.println("Non batch incidents..\n");
			for (int row = 20; row < 26; row++) {
				while (incidentsClosedByEachSet.next() && tasksClosedByEachSet.next()) {
					System.out.println(
							incidentsClosedByEachSet.getString(1) + "\t\t" + incidentsClosedByEachSet.getInt(2)+"\t\t"+
					incidentsClosedByEachSet.getInt(3));
					XSSFCell incidentsClosedBy = sheet.getRow(row).getCell(1);
					XSSFCell incidentsClosedTotalCount = sheet.getRow(row).getCell(2);
					XSSFCell incidentsClosedNonBatchCount = sheet.getRow(row).getCell(3);
					XSSFCell incidentsClosedBatchCount = sheet.getRow(row).getCell(4);
					
					
					incidentsClosedBy.setCellValue(incidentsClosedByEachSet.getString(1));
					incidentsClosedTotalCount.setCellValue(incidentsClosedByEachSet.getInt(2));
					incidentsClosedNonBatchCount.setCellValue(incidentsClosedByEachSet.getInt(3));
					incidentsClosedBatchCount.setCellValue(incidentsClosedByEachSet.getInt(2) - incidentsClosedByEachSet.getInt(3));
					break;
					
				}
			}
			file.close();

			FileOutputStream outFile = new FileOutputStream(
					new File("C:\\Users\\vijsu\\Desktop\\Shift Summary Report\\Data.xlsx"));
			workbook.write(outFile);
			outFile.close();
			System.out.println("Report generated successfully!!");
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

}
