package read.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FilenameFilter;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelImporterInDatabase {

	XSSFRow row;

	private final String url = "jdbc:postgresql://localhost/postgres";
	private final String user = "postgres";
	private final String password = "Shinde@123";
	public static String fileName1 = "Student";
	public String fileName = fileName1;// "Vendor";
	List list = new ArrayList();
	boolean flag = true;
	static String[] fileNameExcel;

	public static void main(String[] args) throws ClassNotFoundException, IOException {
		// TODO Auto-generated method stub

		String folderpath = "Resource\\";
		File file1 = new File(folderpath);
		String absolutePath1 = file1.getAbsolutePath();
		System.out.println(absolutePath1);

		ExcelImporterInDatabase database = new ExcelImporterInDatabase();

		File folder = new File(absolutePath1);
		database.getAllFilesWithCertainExtension(folder, "xlsx");

		String fn = "";
		for (int j = 0; j < fileNameExcel.length; j++) {
			fn = "";
			for (int i = 0; i < fileNameExcel[j].length(); i++) {
				if (fileNameExcel[j].charAt(i) == '.') {
					break;
				} else {
					fn = fn + fileNameExcel[j].charAt(i);
				}

			}

			String url = "Resource\\" + fn + ".xlsx";// Resource
			File file = new File(url);
			if (file.exists()) {
				String absolutePath = file.getAbsolutePath();
				database.readFile(absolutePath, fn);
			} else {
				System.out.println("File not found this location");
			}
		}

	}

	public void readFile(String fileName, String tablename) {
		flag = true;
		FileInputStream fis;
		String val;
		StringBuffer c = new StringBuffer();
		int valint = 0;
		try {
			System.out.println(
					"-------------------------------READING THE SPREADSHEET-------------------------------------");
			fis = new FileInputStream(fileName);
			XSSFWorkbook workbookRead = new XSSFWorkbook(fis);
			XSSFSheet spreadsheetRead = workbookRead.getSheetAt(0);

			Iterator<Row> rowIterator = spreadsheetRead.iterator();
			while (rowIterator.hasNext()) {
				row = (XSSFRow) rowIterator.next();
				Iterator<Cell> cellIterator = row.cellIterator();

				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();
					cell.setCellType(CellType.STRING);
					switch (cell.getCellType()) {
					case STRING:// System.out.print(cell.getStringCellValue()+"\t");
						val = cell.getStringCellValue();
						list.add(val);
						break;
					case NUMERIC:// System.out.print(cell.getNumericCellValue()+"\t");
						valint = (int) cell.getNumericCellValue();

						list.add(valint);
						break;
					// case BOOLEAN:System.out.println(cell.getBooleanCellValue());break;
					}
				}
				if (flag == true) {
					flag = false;
					StringBuffer b = new StringBuffer();

					b.append("CREATE TABLE   IF NOT EXISTS " + tablename + "( " + tablename
							+ "_Id SERIAL  PRIMARY KEY,");

					int len = list.size();

					for (int i = 0; i < list.size(); i++) {
						System.out.println(list.get(i));

						if (len == i + 1) {
							b.append(list.get(i) + " varchar(20));");
							c.append(list.get(i) + ")");
						} else {
							b.append(list.get(i) + " varchar(20),");
							c.append(list.get(i) + ",");
						}

					}
					System.out.println(b);
					System.out.println(c);
					createTable(b);
					list.clear();

				}

				else {
					// StringBuffer s = new StringBuffer();
					StringBuffer b = new StringBuffer();

					b.append("insert into " + tablename + "(");
					b.append(c);
					b.append(" values(");

					int len = list.size();

					for (int i = 0; i < list.size(); i++) {
						if (len == i + 1) {
							b.append("\'" + list.get(i) + "\');");
						} else {
							b.append("\'" + list.get(i) + "\' ,");
						}

				
					}
					System.out.println(b);
					insertTable(b);
					list.clear();

				}

			}

			fis.close();
			c = null;

		} catch (IOException e) {

			System.out.println("Please save & close file");
		}
	}

	private void insertTable(StringBuffer b) {
		// TODO Auto-generated method stub
		String s = b.toString();
		try {
			int result = 0;

			Class.forName("org.postgresql.Driver");
			// System.out.println("connected");
			Connection conn = DriverManager.getConnection(url, user, password);
			Statement stmt = conn.createStatement();

			result = stmt.executeUpdate(s);
			if (result == 1) {
				System.out.println("Record Inserted");
			} else {
				// System.out.println("Table already exists");
			}
			conn.close();

		} catch (SQLException e) {

		} catch (ClassNotFoundException e) {
			// TODO Auto-generated catch block
			System.out.println("Class Not Found");
		}
	}

	private void createTable(StringBuffer b) {
		// TODO Auto-generated method stub
		String s = b.toString();
		try {
			int result = 0;

			Class.forName("org.postgresql.Driver");
			System.out.println("connected");
			Connection conn = DriverManager.getConnection(url, user, password);
			Statement stmt = conn.createStatement();

			result = stmt.executeUpdate(s);
			if (result == 1) {
				System.out.println("Table created");
			} else {
				// System.out.println("Table already exists");
			}
			conn.close();

		} catch (SQLException e) {
			System.out.println("Database connection problem");
		} catch (ClassNotFoundException e) {
			// TODO Auto-generated catch block
			System.out.println("Class Not Found");
		}
	}

	public void getAllFilesWithCertainExtension(File folder, String filterExt) {
		MyExtFilter extFilter = new MyExtFilter(filterExt);
		if (!folder.isDirectory()) {
			System.out.println("Not a folder");
		} else {
			// list out all the file name and filter by the extension
			fileNameExcel = folder.list(extFilter);

			if (fileNameExcel.length == 0) {
				System.out.println("no files end with : " + filterExt);
				return;
			}

			for (int i = 0; i < fileNameExcel.length; i++) {
				System.out.println("File :" + fileNameExcel[i]);
			}
		}
	}

	public class MyExtFilter implements FilenameFilter {

		private String ext;

		public MyExtFilter(String ext) {
			this.ext = ext;
		}

		public boolean accept(File dir, String name) {
			return (name.endsWith(ext));
		}
	}

}
