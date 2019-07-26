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
import java.util.Map.Entry;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.formula.functions.Value;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
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
			// new SecondTuesday().writeDatatoColumn(path,sheetName);

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public static Map<Integer, Integer> fetchDaysCount() {
		Map<Integer, Integer> dayCount = new HashMap<Integer, Integer>();
		Map<Integer, String> excelInputData = new HashMap<Integer, String>();
		String[] inputData = new String[100];
		String[] readDate = new String[100];
		try {

			excelInputData = new SecondTuesday().readAndProcessXLData(path, sheetName);

			for (Map.Entry<Integer, String> entry : excelInputData.entrySet()) {
				// System.err.println(entry.getKey()+" ****************** "+entry.getValue());
				if (entry.getValue().contains("Tuesday +")) {
					String[] onlyDays = entry.getValue().toString().split(" ");
					// System.out.println("Only Day Count: " + onlyDays[3]);
					dayCount.put(entry.getKey(), Integer.parseInt(onlyDays[3]));
				}
			}

			// =====================================================//

			for (Map.Entry<Integer, Integer> day : dayCount.entrySet()) {

				inputData[day.getKey().intValue() - 1] = getSecondTuesdayofMonth(7, 2019, 9, day.getValue().intValue());

			}

			String[] splitstartDateTime = new String[5];
			String[] finalStartDateTime = new String[94];
			List<Object> startDateTime = readXLDateInformation(path, sheetName, "Planned_start_date_local");
			for (int rows = 0; rows <= inputData.length-1; rows++) {

				// System.out.println(startDateTime.get(rows).toString());

				splitstartDateTime = startDateTime.get(rows).toString().split(" ");
				
				String consolidate = dateTimeConverter(startDateTime.get(rows).toString()) + " "
						+ splitstartDateTime[3];
				System.err.println(consolidate);
				finalStartDateTime[rows] = consolidate;
				  
				  
				//  System.err.println(finalStartDateTime[rows]); 
				  xlWriteColumn(path, sheetName,
				  rows,"PlannedDate", finalStartDateTime);
				 

			}

		} catch (Exception e) {
			e.printStackTrace();
		}
		return dayCount;
	}

	public static Date getFirstDateOfMonth(Date date) {
		Calendar cal = Calendar.getInstance();
		cal.setTime(date);
		cal.set(Calendar.DAY_OF_MONTH, cal.getActualMinimum(Calendar.DAY_OF_MONTH));
		return cal.getTime();
	}

	public static final String getSecondTuesdayofMonth(int month, int year, int day, int incrementDay) {
		String dateOutput = null;
		try {
			LocalDate ld = LocalDate.of(year, month, day);
			LocalDate ldNew = ld.plusDays(incrementDay);
			ZoneId defaultZoneId = ZoneId.systemDefault();
			Date date = Date.from(ldNew.atStartOfDay(defaultZoneId).toInstant());
			String datePattern = "M/dd/yyyy";
			SimpleDateFormat simpleDateFormat = new SimpleDateFormat(datePattern);
			dateOutput = simpleDateFormat.format(date);

		} catch (Exception e) {
			e.printStackTrace();
		}

		return dateOutput;
	}

	public static String dateTimeConverter(String dateInput) {
		String formatedDate = null;
		try {
			DateFormat formatter = new SimpleDateFormat("E MMM dd HH:mm:ss Z yyyy");
			Date date = (Date) formatter.parse(dateInput);
			// System.out.println(date);

			Calendar cal = Calendar.getInstance();
			cal.setTime(date);
			formatedDate = (cal.get(Calendar.MONTH)+1) + "/" + cal.get(Calendar.DATE) + "/" + cal.get(Calendar.YEAR);
			// System.out.println("formatedDate : " + formatedDate);
		} catch (Exception e) {
			e.printStackTrace();
		}
		return formatedDate;
	}

	public Map<Integer, String> readAndProcessXLData(String spath, String sheetName) {

		Map<Integer, String> excelData = new HashMap<Integer, String>();
		XSSFWorkbook wb;
		try {

			File file = new File(spath);
			FileInputStream fs = new FileInputStream(file);
			wb = new XSSFWorkbook(fs);
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
			}

		} catch (Exception e) {
			e.printStackTrace();
		}
		return excelData;
	}

	public static void writeDatatoColumn(String spath, String sheetName, int rowNum, String cellValue) {
		XSSFWorkbook wb;
		XSSFCell Cell1;
		XSSFRow row1;
		FileOutputStream fos;
		try {
			// Map<Integer, String> schedule = readAndProcessXLData(spath, sheetName);
			File file = new File(spath);

			wb = new XSSFWorkbook();
			XSSFSheet sheet = wb.createSheet(sheetName);

			row1 = sheet.getRow(rowNum);
			if (row1 == null) {
				row1 = sheet.createRow(rowNum);
			}
			Cell1 = row1.getCell(9);
			if (Cell1 == null) {
				Cell1 = row1.createCell(9);
				Cell1.setCellValue(cellValue);
				fos = new FileOutputStream(file);
				wb.write(fos);
				fos.close();
			}

		} catch (Exception e) {
			e.printStackTrace();

		}
	}

	public static void xlWriteColumn(String xlPath, String sheetName, int rowNum, String columnName,
			String columnData[]) throws Exception {
		int colNum = 0;
		boolean columnFound = false;
		String fileName = xlPath;
		File file = new File(fileName);
		File sameFileName = new File(fileName);
		do {
			Thread.sleep(10);
			if (!file.renameTo(sameFileName))
				System.out.println("Please colose the " + fileName);

		} while (!file.renameTo(sameFileName));

		File outFile = new File(xlPath);

		// FileInputStream myStream1 = new FileInputStream(outFile);
		// XSSFWorkbook myWBook = new XSSFWorkbook (myStream1);
		FileInputStream myStream;
		XSSFWorkbook myWB = null;
		DataFormatter dataFormatter = new DataFormatter();

		do {
			try {
				myStream = new FileInputStream(outFile);
				myWB = new XSSFWorkbook(myStream);
			} catch (Exception e) {
				System.out.println("ignore this is in xlRead -->" + e.getMessage());
				Thread.sleep(100);
			}
		} while (myWB == null);

		XSSFSheet oSheet = myWB.getSheet(sheetName);

		int xCols1 = oSheet.getRow(rowNum).getLastCellNum();
		for (short k = 0; k < xCols1; k++) {
			XSSFCell cell1 = oSheet.getRow(0).getCell(k);
			if (cell1 == null)
				cell1 = oSheet.getRow(0).createCell(k);
			String value1 = cell1.toString();
			if (value1.equalsIgnoreCase(columnName)) {
				columnFound = true;
				colNum = k;
				break;
			}
		}
		XSSFCell cell = null;
		for (int i = 1; i <= columnData.length; i++) {
			if (columnFound) {
				XSSFRow row = oSheet.getRow(i);
				if (row == null)
					row = oSheet.createRow(i);
				cell = row.getCell(colNum);

				if (cell == null)
					cell = row.createCell(colNum);
				cell.setCellType(XSSFCell.CELL_TYPE_STRING);
				cell.setCellValue(columnData[i - 1]);
			} else {

				String cellData = dataFormatter.formatCellValue(cell);
				System.out.println("Given column not found--" + cellData);
			}
		}

		FileOutputStream fOut = new FileOutputStream(outFile);
		myWB.write(fOut);
		fOut.flush();
		fOut.close();

	}

	public static String[] xlReadColumn(String sPath, String sheetName, int rowNum, String columnName)
			throws Exception {
		boolean columnFound = false;
		int columnIndex = 0;

		String fileName = sPath;
		File file = new File(fileName);
		File sameFileName = new File(fileName);
		do {
			Thread.sleep(1);
			if (!file.renameTo(sameFileName))
				System.out.println("Please colose the " + fileName);

		} while (!file.renameTo(sameFileName));

		int xRows1, xCols1;
		String[] xData = {};
		try {
			File myXL = new File(sPath);
			// FileInputStream myStream = new FileInputStream(myXL);
			// XSSFWorkbook myWB = new XSSFWorkbook(myStream);

			FileInputStream myStream;
			XSSFWorkbook myWB = null;

			do {
				try {
					myStream = new FileInputStream(myXL);
					myWB = new XSSFWorkbook(myStream);
				} catch (Exception e) {
					System.out.println("ignore this is in xlRead -->" + e.getMessage());
					Thread.sleep(100);
				}
			} while (myWB == null);

			XSSFSheet mySheet1 = myWB.getSheet(sheetName);
			xRows1 = mySheet1.getLastRowNum() + 1;

			xCols1 = mySheet1.getRow(0).getLastCellNum();

			for (short k = 0; k < xCols1; k++) {
				XSSFCell cell1 = mySheet1.getRow(0).getCell(k);
				if (cell1 == null)
					cell1 = mySheet1.getRow(0).createCell(k);
				// String value = cellToString(cell1);
				String value1 = cell1.toString();
				if (value1.equalsIgnoreCase(columnName)) {
					xData = new String[xRows1 - 1];
					columnFound = true;
					columnIndex = k;
					break;
				}
			}
			if (!columnFound)
				System.out.println("Given column not found--" + columnName);

			if (columnFound)
				for (int i = 1; i < xRows1; i++) {
					XSSFRow row1 = mySheet1.getRow(i);
					if (row1 == null)
						row1 = mySheet1.createRow(i);
					XSSFCell cell1 = row1.getCell(columnIndex);
					if (cell1 == null)
						cell1 = row1.createCell(columnIndex);
					String value = cell1.toString();
					xData[i - 1] = value;
				}
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}

		return xData;
	}

	public static List<Object> readXLDateInformation(String path, String sheetName, String columnName) {
		int i;
		XSSFWorkbook wb;
		XSSFRow row = null;
		XSSFCell cell;
		XSSFCell cellData = null;
		ArrayList<Object> al = new ArrayList<Object>();
		boolean flag = false;
		int rowcount;
		int coulIndex = 0;
		try {
			File file = new File(path);
			FileInputStream fs = new FileInputStream(file);
			wb = new XSSFWorkbook(fs);
			XSSFSheet sheet = wb.getSheet(sheetName);
			rowcount = sheet.getLastRowNum();

			int cellCount = sheet.getRow(0).getLastCellNum();
			for (int j = 0; j < cellCount; j++) {
				cell = sheet.getRow(0).getCell(j);
				if (cell.getStringCellValue().equals(columnName)) {
					flag = true;
					coulIndex = j;
					break;
				}

			}

			for (int x = 1; x < rowcount; x++) {
				if (flag) {
					cellData = sheet.getRow(x).getCell(coulIndex);
					if (cellData != null && cellData.getCellType() == Cell.CELL_TYPE_NUMERIC) {
						Date date = cellData.getDateCellValue();
						al.add(date);
					} else if (cellData != null && cellData.getCellType() == Cell.CELL_TYPE_STRING) {
						al.add(cellData.getStringCellValue());
					} else if (cellData != null && DateUtil.isCellDateFormatted(cellData)) {
						al.add(cellData.getDateCellValue());
					}
				}

			}

		} catch (Exception e) {
			e.printStackTrace();
		}
		return al;
	}

}
