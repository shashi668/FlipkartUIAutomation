package MavenTest.com.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.YearMonth;
import java.time.ZoneId;
import java.time.temporal.TemporalAdjusters;
import java.time.temporal.TemporalField;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.formula.functions.Value;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class SecondTuesday {

	private static Calendar cacheCalendar;
	Excel excel = new Excel();
	static String path = "D:\\Herbalife-GOC\\MavenTesting\\Shashank_Rework.xlsx";
	static String sheetName = "SNOW";

	public static void main(String[] args) {
		try {
			fetchDaysCount();
			new SecondTuesday().writeDatatoColumn(path,sheetName);

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public static Map<Integer, Integer> fetchDaysCount() {
		Map<Integer, Integer> dayCount = new HashMap<Integer, Integer>();
		Map<Integer, String> excelInputData = new HashMap<Integer, String>();
		try {

			excelInputData = new SecondTuesday().readAndProcessXLData(path, sheetName);

			for (Map.Entry<Integer, String> entry : excelInputData.entrySet()) {
				// System.err.println(entry.getKey()+" ****************** "+entry.getValue());
				if (entry.getValue().contains("Tuesday +")) {
					String[] onlyDays = entry.getValue().toString().split(" ");
					System.out.println("Only Day Count:  " + onlyDays[3]);
					dayCount.put(entry.getKey(), Integer.parseInt(onlyDays[3]));
				}
			}
			
			//=====================================================//
			
			for(Map.Entry<Integer, Integer> day : dayCount.entrySet())
			{
				//System.out.println("day.getValue().intValue() : "+day.getValue().intValue());
				System.out.println(getSecondTuesdayofMonth(7, 2019,9,day.getValue().intValue()));
			}
			

		} catch (Exception e) {
			e.printStackTrace();
		}
		return dayCount;
	}
	
	public static Date getFirstDateOfMonth(Date date){
        Calendar cal = Calendar.getInstance();
        cal.setTime(date);
        cal.set(Calendar.DAY_OF_MONTH, cal.getActualMinimum(Calendar.DAY_OF_MONTH));
        return cal.getTime();
    }

	public static final String getSecondTuesdayofMonth(int month, int year,int day,int incrementDay) {

	
		LocalDate ld = LocalDate.of(year, month, day);
		LocalDate ldNew = ld.plusDays(incrementDay);
		ZoneId defaultZoneId = ZoneId.systemDefault();
		Date date = Date.from(ldNew.atStartOfDay(defaultZoneId).toInstant());
		String datePattern = "M/dd/yyyy";
		SimpleDateFormat simpleDateFormat = new SimpleDateFormat(datePattern);
		String dateOutput = simpleDateFormat.format(date);
		return dateOutput;
	}

	public Map<Integer, String> readAndProcessXLData(String spath, String sheetName) {

		
		Map<Integer, String> excelData = new HashMap<Integer, String>();
		try {

			File file = new File(spath);
			FileInputStream fs = new FileInputStream(file);
			XSSFWorkbook wb = new XSSFWorkbook(fs);
			XSSFSheet sheet = wb.getSheet(sheetName);
			DataFormatter dataFormatter = new DataFormatter();
			int rowCount = sheet.getLastRowNum();
			for (int i = 1; i <= rowCount; i++) {
				// System.err.println("rowCount " + i);

				XSSFRow row = sheet.getRow(i);
				int cellCount = row.getLastCellNum();
				// for (int j = 0; j < cellCount; j++) {
				XSSFCell cell = row.getCell(4);
				String cellData = dataFormatter.formatCellValue(cell);
				excelData.put(i, cellData);

				// }
			}

		} catch (Exception e) {
			e.printStackTrace();
		}
		return excelData;
	}
	
	
	public void writeDatatoColumn(String spath,String sheetName)
	{
		try
		{
			Map<Integer, String> schedule = readAndProcessXLData(spath, sheetName);
			File file = new File(spath);
			
			XSSFWorkbook wb = new XSSFWorkbook();
			XSSFSheet sheet = wb.createSheet(sheetName);
			System.err.println(schedule.entrySet());
			
			FileOutputStream  fos = new FileOutputStream(file);
			wb.write(fos);
			fos.close();
		}catch (Exception e) {
			e.printStackTrace();
			
		}
	}
	

}
