package utility;

import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Objects;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.Reporter;

import helper.AssertionHelper;

public class ExcelReader {

	private static volatile XSSFSheet ExcelWSheet;
	private static volatile XSSFWorkbook ExcelWBook;
	private static XSSFCell Cell;
	ConfigFileReader configFileReader = new ConfigFileReader();

	public XSSFSheet getSheet(String ExcelSheetName, String SheetName) throws Exception {
		HashMap<String, XSSFSheet> sheetFromExcel = new HashMap<>();
		try {
			// Open the Excel file
			FileInputStream ExcelFile = new FileInputStream(configFileReader.getValue(ExcelSheetName.toUpperCase()));
			// Access the required test data sheet
			ExcelWBook = new XSSFWorkbook(ExcelFile);
			for (int i = 0; i < ExcelWBook.getNumberOfSheets(); i++) {
				sheetFromExcel.put(ExcelWBook.getSheetAt(i).getSheetName(), ExcelWSheet = ExcelWBook.getSheetAt(i));
			}
		} catch (Exception e) {
			throw (e);
		}
		for (Entry<String, XSSFSheet> entry : sheetFromExcel.entrySet()) {
			if (SheetName.equals(entry.getKey())) {
				ExcelWSheet = entry.getValue();
			}
		}
		return ExcelWSheet;
	}

	public int gettoatalrow(String ExcelSheetName, String SheetName) throws Exception {
		HashMap<String, XSSFSheet> sheetFromExcel = new HashMap<>();
		int size = 0;
		try {
			// Open the Excel file
			FileInputStream ExcelFile = new FileInputStream(configFileReader.getValue(ExcelSheetName.toUpperCase()));
			// Access the required test data sheet
			ExcelWBook = new XSSFWorkbook(ExcelFile);
			for (int i = 0; i < ExcelWBook.getNumberOfSheets(); i++) {
				sheetFromExcel.put(ExcelWBook.getSheetAt(i).getSheetName(), ExcelWSheet = ExcelWBook.getSheetAt(i));
			}
		} catch (Exception e) {
			throw (e);
		}
		for (Entry<String, XSSFSheet> entry : sheetFromExcel.entrySet()) {
			if (SheetName.equals(entry.getKey())) {
				ExcelWSheet = entry.getValue();
				try {

					for (int g = 0; g <= ExcelWSheet.getPhysicalNumberOfRows(); g++) {
						String ct = ExcelWSheet.getRow(g).getCell(0).getStringCellValue();
						char c = ct.charAt(0);

						if ((c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z')) {
							size = size + 1;
						}
					}
				} catch (Exception e) {
					Reporter.log(e.getMessage());
				}
			}
		}

		return size - 1;
	}

	public String getValueFromExcel(HashMap<String, String> totHash, String ColumnName) {
		String columnValue = null;
		if (totHash.size() != 0) {
			ArrayList<String> hh = getValue(totHash, ColumnName);
			if (hh.size() == 1) {
				columnValue = hh.toString();
			} else {
				AssertionHelper.markFail("Duplicate value Preset:" + hh.toString());
			}
		} else {
			AssertionHelper.markFail("The Excel To hash is Null");
		}
		if (columnValue.contains("[") && columnValue.contains("]")) {
			columnValue = columnValue.replace("[", "").replace("]", "");
		}
		return columnValue;
	}

	public HashMap<String, String> getHashValueFromExcel(XSSFSheet ExcelWSheet, String TestCaseId, String SerialNo) {
		HashMap<String, String> totHash = new LinkedHashMap<>();

		try {

			// needs to change code here to get only row value instead of whole
			// sheet
			for (int i = 0; i <= ExcelWSheet.getLastRowNum(); i++) {
				if (TestCaseId.equals(ExcelWSheet.getRow(i).getCell(0).getStringCellValue())) {
					if (SerialNo.equals(ExcelWSheet.getRow(i).getCell(1).getStringCellValue())) {
						for (int j = 0; j <= ExcelWSheet.getRow(i).getLastCellNum() - 1; j++) {
							String KeyForMap = "";
							String valueForMap = "";
							KeyForMap = ExcelWSheet.getRow(0).getCell(j).getStringCellValue();
							valueForMap = ExcelWSheet.getRow(i).getCell(j).getStringCellValue();

							totHash.put(KeyForMap.toUpperCase(), valueForMap);

						}
					}
				}

			}

		} catch (Exception e) {

		}
		return totHash;
	}

	public HashMap<String, String> getHashValueFromExcelWithCOLOUR(XSSFSheet ExcelWSheet, String TestCaseId,
			String SerialNo) {
		HashMap<String, String> totHash = new LinkedHashMap<>();

		try {

			// needs to change code here to get only row value instead of whole
			// sheet
			for (int i = 0; i <= ExcelWSheet.getLastRowNum(); i++) {
				if (TestCaseId.equals(ExcelWSheet.getRow(i).getCell(0).getStringCellValue())) {
					if (SerialNo.equals(ExcelWSheet.getRow(i).getCell(1).getStringCellValue())) {
						for (int j = 0; j <= ExcelWSheet.getRow(i).getLastCellNum() - 1; j++) {
							String KeyForMap = "";
							String valueForMap = "";
							KeyForMap = ExcelWSheet.getRow(0).getCell(j).getStringCellValue() + "_"
									+ (ExcelWSheet.getRow(0).getCell(j).getCellStyle().getFillForegroundColorColor())
											.getARGBHex();
							valueForMap = ExcelWSheet.getRow(i).getCell(j).getStringCellValue();

							totHash.put(KeyForMap.toUpperCase(), valueForMap);

						}
					}
				}

			}

		} catch (Exception e) {

		}
		totHash.values().removeIf(Objects::isNull);
		return totHash;
	}

	public HashMap<String, String> getHashValueFromExcel(XSSFSheet ExcelWSheet, String TestCaseId) {
		HashMap<String, String> totHash = new LinkedHashMap<>();
		try {

			// needs to change code here to get only row value instead of whole
			// sheet
			for (int i = 0; i <= ExcelWSheet.getLastRowNum(); i++) {
				if (TestCaseId.equals(ExcelWSheet.getRow(i).getCell(0).getStringCellValue())) {

					for (int j = 0; j <= ExcelWSheet.getRow(i).getLastCellNum() - 1; j++) {
						String KeyForMap = "";
						String valueForMap = "";
						KeyForMap = ExcelWSheet.getRow(0).getCell(j).getStringCellValue();
						valueForMap = ExcelWSheet.getRow(i).getCell(j).getStringCellValue();

						totHash.put(KeyForMap.toUpperCase(), valueForMap);

					}
				}

			}

		} catch (Exception e) {

		}
		totHash.values().removeIf(Objects::isNull);
		return totHash;
	}

	public HashMap<String, String> getHashValueFromExcelForCaseSensitvie(XSSFSheet ExcelWSheet, String TestCaseId) {
		HashMap<String, String> totHash = new LinkedHashMap<>();
		try {

			// needs to change code here to get only row value instead of whole
			// sheet
			for (int i = 0; i <= ExcelWSheet.getLastRowNum(); i++) {
				if (TestCaseId.equals(ExcelWSheet.getRow(i).getCell(0).getStringCellValue())) {

					for (int j = 0; j <= ExcelWSheet.getRow(i).getLastCellNum() - 1; j++) {
						String KeyForMap = "";
						String valueForMap = "";
						KeyForMap = ExcelWSheet.getRow(0).getCell(j).getStringCellValue().toUpperCase();
						valueForMap = ExcelWSheet.getRow(i).getCell(j).getStringCellValue();

						totHash.put(KeyForMap, valueForMap);

					}
				}

			}

		} catch (Exception e) {

		}
		totHash.values().removeIf(Objects::isNull);
		return totHash;
	}

	public void ceckTotalTestDataWithDB(HashMap<String, String> DBHsh, HashMap<String, String> ExcelCreDataHsh) {
		for (Map.Entry<String, String> entryCRE : DBHsh.entrySet()) {
			if (ExcelCreDataHsh.containsKey(entryCRE.getKey())) {
				if (!"NULL".equals(ExcelCreDataHsh.get(entryCRE.getKey().trim()).toUpperCase())) {
					if (entryCRE.getValue().trim().equals(ExcelCreDataHsh.get(entryCRE.getKey().trim()))) {
						AssertionHelper.markPass("PASS:CRE VALUE :" + entryCRE.getKey() + ":" + entryCRE.getValue()
								+ "Validated With  " + ExcelCreDataHsh.get(entryCRE.getKey()) + "  PASS");

					} else {
						AssertionHelper.markPass("FAIL:CRE VALUE :" + entryCRE.getKey() + ":" + entryCRE.getValue()
								+ "Validated With  " + ExcelCreDataHsh.get(entryCRE.getKey()) + "  FAIL");

					}
				} else {
					AssertionHelper.markPass("PASS:CRE VALUE :" + entryCRE.getValue() + "Validated With  NULL"
							+ ExcelCreDataHsh.get(entryCRE.getKey()) + "PASS:  VAlue is NUll");

				}
			}
		}

	}

	public HashMap<String, String> getHashValueFromExcel(XSSFSheet ExcelWSheet) {
		HashMap<String, String> totHash = new LinkedHashMap<>();
		try {

			// needs to change code here to get only row value instead of whole
			// sheet
			for (int i = 0; i <= ExcelWSheet.getLastRowNum(); i++) {

				for (int j = 0; j <= ExcelWSheet.getRow(i).getLastCellNum() - 1; j++) {
					String KeyForMap = "";
					String valueForMap = "";
					KeyForMap = ExcelWSheet.getRow(0).getCell(j).getStringCellValue();
					valueForMap = ExcelWSheet.getRow(i).getCell(j).getStringCellValue();

					totHash.put(KeyForMap.toUpperCase(), valueForMap);

				}

			}
		} finally {

		}
		return totHash;
	}

	// This method is to read the test data from the Excel cell, in this we are
	// passing parameters as Row num and Col num

	public static String getCellData(int RowNum, int ColNum) {
		try {
			Cell = ExcelWSheet.getRow(RowNum).getCell(ColNum);
			String CellData = Cell.getStringCellValue();
			return CellData;
		} catch (Exception e) {
			return "";
		}
	}

	public static String getTestCaseName(String sTestCase) throws Exception {
		String value = sTestCase;
		try {
			int posi = value.indexOf('@');
			value = value.substring(0, posi);
			posi = value.lastIndexOf('.');
			value = value.substring(posi + 1);
			return value;
		} catch (Exception e) {
			throw (e);
		}
	}

	public static int getRowContains(String sTestCaseName, int colNum) throws Exception {
		int i;
		try {
			int rowCount = ExcelReader.getRowUsed();
			for (i = 0; i < rowCount; i++) {
				if (ExcelReader.getCellData(i, colNum).equalsIgnoreCase(sTestCaseName)) {
					break;
				}
			}
			return i;
		} catch (Exception e) {
			throw (e);
		}
	}

	public String[] getdataProvidercount(String ExcelSheetName, String SheetName) {
		int val = 0;

		try {
			val = gettoatalrow(ExcelSheetName, SheetName);
		} catch (Exception e) {
			Reporter.log(e.getMessage());
		}
		String arr[] = new String[val];

		for (int v = 0; v <= val - 1; v++) {
			arr[v] = "SL" + (v + 1);
		}

		return arr;

	}

	public static int getRowUsed() throws Exception {
		try {
			int RowCount = ExcelWSheet.getLastRowNum();
			return RowCount;
		} catch (Exception e) {

			throw (e);
		}
	}

	public static <K, V> ArrayList<String> getKey(Map<K, V> map, V value) {
		ArrayList<String> blist = new ArrayList<String>();
		for (Map.Entry<K, V> entry : map.entrySet()) {
			if (value.equals(entry.getValue())) {
				blist.add(entry.getKey().toString());
			}
		}
		return blist;
	}

	public static <K, V> ArrayList<String> getValue(Map<K, V> map, V value) {
		ArrayList<String> blist = new ArrayList<String>();
		for (Map.Entry<K, V> entry : map.entrySet()) {
			if (value.equals(entry.getKey())) {
				blist.add(entry.getValue().toString());
			}
		}
		return blist;
	}

	public HashMap<String, String> getSheetDataFromExcel(String ExcelSheetName, String SheetName, String TestCaseId,
			String slno) throws Exception {
		HashMap<String, String> sheetDataFromExcel = new HashMap<>();
		XSSFSheet sheet;
		sheet = getSheet(ExcelSheetName, SheetName);
		sheetDataFromExcel = getHashValueFromExcel(sheet, TestCaseId, slno);
		return sheetDataFromExcel;

	}

	public HashMap<String, String> getSheetDataFromExcel(String ExcelSheetName, String SheetName, String TestCaseId)
			throws Exception {
		HashMap<String, String> sheetDataFromExcel = new HashMap<>();
		XSSFSheet sheet;
		sheet = getSheet(ExcelSheetName, SheetName);
		sheetDataFromExcel = getHashValueFromExcel(sheet, TestCaseId);
		return sheetDataFromExcel;
	}

}
