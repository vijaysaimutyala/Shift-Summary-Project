import java.sql.*;
import java.io.*;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.record.cf.Threshold;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
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
	static Float inflowOutflowRatio, memberAverageResolvedBatch, memberAverageResolvedNonBatch,
			wipPercentage,toBeCheckedFloat,outflowFloat,inflowFloat;
	static int toBeChecked, withAdvTeam, withOtherTeams, withUsers, incidentsClosedByEach, tasksClosedByEach, inflow, outflow,
	totalIncidentsClosed,totalBatchIncidentsClosed,totalNonBatchIncidentsClosed;
	static ResultSet countFromPrevShift, countHandedOverToNextShift, toBeCheckedSet, withAdvTeamSet, withUsersSet,
			withOtherTeamsSet, tasksClosedByEachSet, batchTasksClosedByEachSet, incidentsClosedByEachSet,
			batchIncidentsClosedByEachSet;
	static XSSFCellStyle thresholdValuesStyle;
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
			
			inflowFloat = 0.0f;
			while (countFromPrevShift.next()) {
				System.out.println("Count From Previous Shift: " + countFromPrevShift.getFloat(1));
				inflow = inflow + countFromPrevShift.getInt(1);
			}
			outflowFloat = 0.0f;
			while (countHandedOverToNextShift.next()) {
				System.out.println("Count From Handedover to Next Shift: " + countHandedOverToNextShift.getFloat(1)+"\n------------------------------------------------");
				outflow =outflow +  countHandedOverToNextShift.getInt(1);
				outflowFloat = outflowFloat + countHandedOverToNextShift.getFloat(1);
			}
			System.out.println("To Be Checked \n Location \t\t Count ");
			toBeCheckedFloat = 0.0f;
			while (toBeCheckedSet.next()) {

				System.out.println(toBeCheckedSet.getString(1) + "\t\t" + toBeCheckedSet.getString(2)+"\n------------------------------------------------");
				toBeChecked  = toBeChecked + toBeCheckedSet.getInt(2);
				toBeCheckedFloat = toBeCheckedFloat + toBeCheckedSet.getFloat(2);
			}
			System.out.println("To be Checked is " + toBeChecked);

			System.out.println("With Adv Teams \n Location \t\t Count ");
			withAdvTeam = 0;
			while (withAdvTeamSet.next()) {
				System.out.println(withAdvTeamSet.getString(1) + "\t\t" + withAdvTeamSet.getInt(2)+"\n------------------------------------------------");
				withAdvTeam = withAdvTeam + withAdvTeamSet.getInt(2);
			}
			System.out.println("With Adv Team is "+withAdvTeam);
			System.out.println("With Users \n Location \t\t Count ");
			withUsers = 0;
			while (withUsersSet.next()) {
				System.out.println(withUsersSet.getString(1) + "\t\t" + withUsersSet.getInt(2)+"\n------------------------------------------------");
				withUsers = withUsers + withUsersSet.getInt(2);
			}
			System.out.println("With Users is "+withUsers);

			System.out.println("With Other Teams \n Location \t\t Count ");
			while (withOtherTeamsSet.next()) {
				System.out.println(withOtherTeamsSet.getString(1) + "\t\t" + withOtherTeamsSet.getInt(2)+"\n------------------------------------------------");
				withOtherTeams = withOtherTeams + withOtherTeamsSet.getInt(2);
			}
			System.out.println("With Other Teams is "+withOtherTeams);
			System.out.println("Tasks closed by each \n Closed By \t\t Count ");
			while (tasksClosedByEachSet.next()) {
				System.out.println(tasksClosedByEachSet.getString(1) + "\t\t" + tasksClosedByEachSet.getInt(2)+"\n------------------------------------------------");
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

			inflowOutflowRatio = (inflowFloat / outflowFloat);
			wipPercentage =  (toBeCheckedFloat / outflowFloat)*100;
			System.out.println("WIP % "+wipPercentage);
			System.out.println( "To be checked "+toBeChecked);
			System.out.println("Outflow "+outflow);
			System.out.println("Inflow / Outflow Ratio : \t" + inflowOutflowRatio);
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
			thresholdValuesStyle = workbook.createCellStyle();
			
			// ------Update the Carry Forward From Previous Shift------//
			XSSFCell carryForwardFromPreviousShift = sheet.getRow(2).getCell(2);
			carryForwardFromPreviousShift.setCellValue(inflow);
			
			//-------------Update WIP %------------------------//
			XSSFCell wipPercentageCount = sheet.getRow(10).getCell(2);
			wipPercentageCount.setCellValue(wipPercentage);
			if(wipPercentage<=20){
				thresholdValuesStyle.setFillForegroundColor(IndexedColors.GREEN.getIndex());
				thresholdValuesStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
				wipPercentageCount.setCellStyle(thresholdValuesStyle);
			}else{
				thresholdValuesStyle.setFillForegroundColor(IndexedColors.ORANGE.getIndex());
				thresholdValuesStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
				wipPercentageCount.setCellStyle(thresholdValuesStyle);
			}
			
			//----------Update Inflow/Outflow Ratio------------//
			XSSFCell inflowOutflowRatioCount = sheet.getRow(9).getCell(2);
			inflowOutflowRatioCount.setCellValue(inflowOutflowRatio);

			if(inflowOutflowRatio >= 1){
				thresholdValuesStyle.setFillForegroundColor(IndexedColors.GREEN.getIndex());
				thresholdValuesStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
				inflowOutflowRatioCount.setCellStyle(thresholdValuesStyle);
			}else{
				thresholdValuesStyle.setFillForegroundColor(IndexedColors.ORANGE.getIndex());
				thresholdValuesStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
				inflowOutflowRatioCount.setCellStyle(thresholdValuesStyle);
			}
		
					//--------------To be checked Count----------------------//
			XSSFCell toBeCheckedCount = sheet.getRow(13).getCell(6);
			toBeCheckedCount.setCellValue(toBeChecked);	//------With other teams and to be monitored Count---------//
			XSSFCell withOtherTeamsCount = sheet.getRow(14).getCell(6);
			withOtherTeamsCount.setCellValue(withOtherTeams);	
			//-----Update with Users Count------------------------//
			XSSFCell withUsersCount = sheet.getRow(15).getCell(6);
			withUsersCount.setCellValue(withUsers);
			//------Update with Adv Team Count--------------------//
			XSSFCell withAdvTeamCount = sheet.getRow(16).getCell(6);
			withAdvTeamCount.setCellValue(withAdvTeam);
	
		
	
			//------Total passed on count updated in table--------//
			XSSFCell totalCountPassedOn = sheet.getRow(17).getCell(6);
			totalCountPassedOn.setCellValue(outflow);

			
			
			// ---Update Incidents closed by each member----//
			System.out.println("Non batch incidents..\n");
			XSSFCell totalIncidentClosedCount = sheet.getRow(26).getCell(2);
			XSSFCell totalNonBatchIncidentsClosedCount = sheet.getRow(26).getCell(3);
			XSSFCell totalBatchIncidentsClosedCount = sheet.getRow(26).getCell(4);
			for (int row = 20; row < 26; row++) {
				while (incidentsClosedByEachSet.next()) {
					System.out.println(
							incidentsClosedByEachSet.getString(1) + "\t\t" + incidentsClosedByEachSet.getInt(2)+"\t\t"+
					incidentsClosedByEachSet.getInt(3));
					XSSFCell incidentsClosedBy = sheet.getRow(row).getCell(1);
					XSSFCell incidentsClosedTotalCount = sheet.getRow(row).getCell(2);
					XSSFCell incidentsClosedNonBatchCount = sheet.getRow(row).getCell(3);
					XSSFCell incidentsClosedBatchCount = sheet.getRow(row).getCell(4);
										
					totalIncidentsClosed = totalIncidentsClosed + incidentsClosedByEachSet.getInt(2);
					totalNonBatchIncidentsClosed = totalNonBatchIncidentsClosed + incidentsClosedByEachSet.getInt(3);
					totalBatchIncidentsClosed = totalBatchIncidentsClosed + (incidentsClosedByEachSet.getInt(2) - incidentsClosedByEachSet.getInt(3));
					
					
					incidentsClosedBy.setCellValue(incidentsClosedByEachSet.getString(1));
					incidentsClosedTotalCount.setCellValue(incidentsClosedByEachSet.getInt(2));
					incidentsClosedNonBatchCount.setCellValue(incidentsClosedByEachSet.getInt(3));
					incidentsClosedBatchCount.setCellValue(incidentsClosedByEachSet.getInt(2) - incidentsClosedByEachSet.getInt(3));
					break;
					
				}
			}
			totalIncidentClosedCount.setCellValue(totalIncidentsClosed);
			totalNonBatchIncidentsClosedCount.setCellValue(totalNonBatchIncidentsClosed);
			totalBatchIncidentsClosedCount.setCellValue(totalBatchIncidentsClosed);
			
			
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
