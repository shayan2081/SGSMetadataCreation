import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Properties;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DatasetComparion {

	static String DatasetInformationFile; 
	
	public static void main(String[] args) {

		try {

			Properties prop = new Properties();
			String fileName = "\\src\\config.properties";
			InputStream is = new FileInputStream(System.getProperty("user.dir") + fileName);
			prop.load(is);
			String DownloadsPath = (String) prop.get("DownloadsPath");
			DatasetInformationFile = (String) prop.getProperty("DatasetInformationFile");

			List<String> DS250 = new ArrayList<String>();
			List<String> DS500 = new ArrayList<String>();
			List<String> DS1000 = new ArrayList<String>();
			List<String> DS2000 = new ArrayList<String>();
			List<String> DSMODS = new ArrayList<String>();
			List<String> DSSAMPLES = new ArrayList<String>();
			List<String> DSBORHOLES = new ArrayList<String>();

			// int i=0;

			File file = new File(DownloadsPath + DatasetInformationFile);
			FileInputStream fis = new FileInputStream(file);
			XSSFWorkbook wb = new XSSFWorkbook(fis);
			Row row;
			//////////////////////////////////////////////////
			XSSFSheet sheet = wb.getSheetAt(1);
			Iterator<Row> itr = sheet.iterator();
			itr.next();
			while (itr.hasNext()) {
				row = itr.next();
				DS250.add(row.getCell(4).getStringCellValue());
			}
			///////////////////////////////////////////////////
			sheet = wb.getSheetAt(2);
			itr = sheet.iterator();
			itr.next();
			while (itr.hasNext()) {
				row = itr.next();
				DS500.add(row.getCell(1).getStringCellValue());
			}
			///////////////////////////////////////////////////
			sheet = wb.getSheetAt(3);
			itr = sheet.iterator();
			itr.next();
			while (itr.hasNext()) {
				row = itr.next();
				DS1000.add(row.getCell(1).getStringCellValue());
			}
			///////////////////////////////////////////////////
			sheet = wb.getSheetAt(4);
			itr = sheet.iterator();
			itr.next();
			while (itr.hasNext()) {
				row = itr.next();
				DS2000.add(row.getCell(1).getStringCellValue());
			}
			///////////////////////////////////////////////////
			sheet = wb.getSheetAt(6);
			itr = sheet.iterator();
			itr.next();
			while (itr.hasNext()) {
				row = itr.next();
				DSMODS.add(row.getCell(1).getStringCellValue());
			}
			///////////////////////////////////////////////////
			sheet = wb.getSheetAt(7);
			itr = sheet.iterator();
			itr.next();
			while (itr.hasNext()) {
				row = itr.next();
				DSSAMPLES.add(row.getCell(1).getStringCellValue());
			}
			///////////////////////////////////////////////////
			sheet = wb.getSheetAt(8);
			itr = sheet.iterator();
			itr.next();
			while (itr.hasNext()) {
				row = itr.next();
				DSBORHOLES.add(row.getCell(1).getStringCellValue());
			}

			wb.close();

			file = new File(DownloadsPath + DatasetInformationFile);
			fis = new FileInputStream(file);
			wb = new XSSFWorkbook(fis);
			sheet = wb.getSheetAt(0);
			itr = sheet.iterator();
			itr.next();
			while (itr.hasNext()) {

				row = itr.next();
				System.out.println("Dataset ID: "+row.getCell(0).getStringCellValue());

				if (DS250.contains(row.getCell(0).getStringCellValue()))
					row.getCell(15).setCellValue("true");
				else
					row.getCell(15).setCellValue("false");

				if (DS500.contains(row.getCell(0).getStringCellValue()))
					row.getCell(16).setCellValue("true");
				else
					row.getCell(16).setCellValue("false");

				if (DS1000.contains(row.getCell(0).getStringCellValue()))
					row.getCell(17).setCellValue("true");
				else
					row.getCell(17).setCellValue("false");

				if (DS2000.contains(row.getCell(0).getStringCellValue()))
					row.getCell(18).setCellValue("true");
				else
					row.getCell(18).setCellValue("false");

				if (DSMODS.contains(row.getCell(0).getStringCellValue()))
					row.getCell(19).setCellValue("true");
				else
					row.getCell(19).setCellValue("false");

				if (DSSAMPLES.contains(row.getCell(0).getStringCellValue()))
					row.getCell(20).setCellValue("true");
				else
					row.getCell(20).setCellValue("false");

				if (DSBORHOLES.contains(row.getCell(0).getStringCellValue()))
					row.getCell(21).setCellValue("true");
				else
					row.getCell(21).setCellValue("false");
			}

			fis.close();

			FileOutputStream outputStream = new FileOutputStream(DownloadsPath + DatasetInformationFile);
			wb.write(outputStream);
			wb.close();
			outputStream.close();

			wb.close();

		} catch (Exception e) {
			//
			e.printStackTrace();
		}

	}

}
