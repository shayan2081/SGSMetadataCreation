
import java.io.File;

import java.io.FileInputStream;

import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;

import java.io.PrintStream;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;

import java.util.Iterator;

import java.util.List;
import java.util.Properties;

import org.apache.poi.ss.usermodel.Cell;

import org.apache.poi.ss.usermodel.CellType;

import org.apache.poi.ss.usermodel.Row;

import org.apache.poi.xssf.usermodel.XSSFSheet;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ArabianPlatform1000K {

	static String DownloadsPath;
	static String GeneralInformationFile;
	static String DatasetInformationFile;
	static String DataFieldsDescriptionFile; 

	public static void main(String[] args) {

		try

		{

			Properties prop = new Properties();
			String fileName = "\\src\\config.properties";
			InputStream is = new FileInputStream(System.getProperty("user.dir") + fileName);
			prop.load(is);
			DownloadsPath = (String) prop.get("DownloadsPath");
			GeneralInformationFile = (String) prop.getProperty("GeneralInformationFile");
			DatasetInformationFile = (String) prop.getProperty("DatasetInformationFile");
			DataFieldsDescriptionFile = (String) prop.getProperty("DataFieldsDescriptionFile");
			
			
			System.out.println("General Information - starting \n\n");

			List<String> GeneralDatasetColumnsFirstRow = new ArrayList<String>();

			List<String> GeneralDatasetColumnsSecondRow = new ArrayList<String>();

			List<String> GeneralDatasetValues = new ArrayList<String>();

			File file = new File(DownloadsPath + GeneralInformationFile);

			FileInputStream fis = new FileInputStream(file);

			XSSFWorkbook wb = new XSSFWorkbook(fis);

			XSSFSheet sheet = wb.getSheetAt(0);

			Iterator<Row> itr = sheet.iterator();

			Row row = sheet.getRow(0); // First Row

			Iterator<Cell> cellIterator = row.cellIterator();

			int total = 25;

			int count = 0;

			while (count < total) {

				if (row.getCell(count) != null)

					GeneralDatasetColumnsFirstRow.add("" + row.getCell(count).getStringCellValue());

				else

					GeneralDatasetColumnsFirstRow.add("");

				count++;

			}

			row = sheet.getRow(1); // Second Row

			count = 0;

			while (count < total) {

				if (row.getCell(count) != null)

					GeneralDatasetColumnsSecondRow.add("" + row.getCell(count).getStringCellValue());

				else

					GeneralDatasetColumnsSecondRow.add("");

				count++;

			}

			row = sheet.getRow(2); // Third Row

			count = 0;

			while (count < total) {

				if (row.getCell(count) != null)

					GeneralDatasetValues.add("" + row.getCell(count).getStringCellValue());

				else

					GeneralDatasetValues.add("");

				count++;

			}

			StringBuilder GeneralDatasetInformation = new StringBuilder("");

			for (int i = 0; i < total; i++) {

				if (GeneralDatasetValues.get(i) != "") {
					if (GeneralDatasetColumnsSecondRow.get(i).equals("Access constraints")) {

						GeneralDatasetInformation.append(
								"<div class=\"propertyLabel\"><span class=\"propertyLabelSpan\">Package Version</span></div>\r\n"

										+ "<div class=\"propertyValue\">1.0</div>\r\n"

										+ "<div class=\"propertyLabel\"><span class=\"propertyLabelSpan\">Package Date</span></div>\r\n"

										+ "<div class=\"propertyValue\">31/10/2021</div>\r\n");

					}

					if (GeneralDatasetColumnsSecondRow.get(i).equals(""))

						GeneralDatasetInformation
								.append("<div class=\"propertyLabel\"><span class=\"propertyLabelSpan\">"
										+ GeneralDatasetColumnsFirstRow.get(i) + "</span></div>\r\n"

										+ "<div class=\"propertyValue\">" + GeneralDatasetValues.get(i) + "</div>\r\n");

					else

						GeneralDatasetInformation
								.append("<div class=\"propertyLabel\"><span class=\"propertyLabelSpan\">"
										+ GeneralDatasetColumnsSecondRow.get(i) + "</span></div>\r\n"

										+ "<div class=\"propertyValue\">" + GeneralDatasetValues.get(i) + "</div>\r\n");
				}

			}

			itr.next();

			wb.close();
			
			System.out.println("General Information - finished \n\n");

			/////////////////////////////////////////////////////////////////////

			System.out.println("Data Fields Description - starting \n\n");
			
			StringBuilder DataFieldsDescriptionTable = new StringBuilder();

			file = new File(DownloadsPath + DataFieldsDescriptionFile);

			fis = new FileInputStream(file);

			wb = new XSSFWorkbook(fis);

			sheet = wb.getSheetAt(1);

			itr = sheet.iterator();

			itr.next(); // Leave 1st line

			while (itr.hasNext())

			{

				row = itr.next();

				cellIterator = row.cellIterator();

				DataFieldsDescriptionTable.append("<tr>");

				while (cellIterator.hasNext())

				{

					DataFieldsDescriptionTable.append("<td>");

					DataFieldsDescriptionTable.append(cellIterator.next().getStringCellValue());

					DataFieldsDescriptionTable.append("</td>");

				}

				DataFieldsDescriptionTable.append("</tr>");

			}

			wb.close();
			
			System.out.println("Data Fields Description - finished \n\n");

			/////////////////////////////////////////////////////////////////////

			System.out.println("Dataset Information - starting \n\n");
			
			file = new File(DownloadsPath + DatasetInformationFile);

			fis = new FileInputStream(file);

			wb = new XSSFWorkbook(fis);

			sheet = wb.getSheetAt(3);

			itr = sheet.iterator();

			row = itr.next(); // first row (Headings Column)

			List<String> DatasetInformationFirstRow = new ArrayList<String>();

			cellIterator = row.cellIterator();
			total = row.getPhysicalNumberOfCells();
			count = 0;
			while (count < total) {
				if (row.getCell(count) != null)
					DatasetInformationFirstRow.add("" + row.getCell(count).getStringCellValue());
				else
					DatasetInformationFirstRow.add("");
				count++;
			}
			
			System.out.println("Dataset Information - finished \n\n");
			
			System.out.println("Creating HTML Files - starting\n\n");

			while (itr.hasNext()) {
				row = itr.next();
				createHtmlFile(row, DatasetInformationFirstRow, GeneralDatasetInformation, DataFieldsDescriptionTable);
			}

			wb.close();
			
			System.out.println("\nCreating HTML Files - finished \n\n");

		}

		catch (Exception e)

		{

			e.printStackTrace();

		}

	}

	public static void createHtmlFile(Row row, List<String> DatasetInformationFirstRow,
			StringBuilder GeneralDatasetInformation, StringBuilder DataFieldsDescriptionTable)

	{

		Iterator<Cell> cellIterator = row.cellIterator();

		List<String> columns = new ArrayList<String>();

		while (cellIterator.hasNext())

		{

			Cell cell = cellIterator.next();

			if (cell.getCellType() == CellType.NUMERIC)

				columns.add(String.valueOf((int) cell.getNumericCellValue()));

			else

				columns.add(cell.getStringCellValue());

		}

		try {

			System.out.println("DatasetID:"+columns.get(0));
			
			File newFile = new File("c:\\Html Files\\DS_1000K\\" + columns.get(0));
			newFile.mkdirs();
			OutputStream htmlfile = new FileOutputStream(
					new File("c:\\Html Files\\DS_1000K\\" + columns.get(0) + "\\" + columns.get(0) + "_Metadata.html"));

			File wordfile = new File(DownloadsPath + "SGS-NGD DATA PRIVACY AGREEMENT.pdf");

			Files.copy(wordfile.toPath(),
					new File("c:\\Html Files\\DS_1000K\\" + columns.get(0) + "\\SGS-NGD DATA PRIVACY AGREEMENT.pdf")
							.toPath(),
					StandardCopyOption.REPLACE_EXISTING);

			PrintStream printhtml = new PrintStream(htmlfile);

			String htmlheader = "<html><head>";

			htmlheader += "<title>Metadata</title> <style type=\"text/css\">\r\n"

					+ "\r\n"

					+ "               \r\n"

					+ "\r\n"

					+ "* { font-family: verdana,arial,sans-serif; font-size: 11px; line-height:150%; }\r\n"

					+ "          h1 { font-size:16px; }\r\n"

					+ "          .header { background-color: #1C9495; color:#ffffff; font-weight: bold; }\r\n"

					+ "          .body { background-color: #FFFFFF; font-family: verdana,arial,sans-serif; font-size: 11px; line-height:150%; }\r\n"

					+ "          .title { text-align: left; min-width:100px; padding-left:15px; }\r\n"

					+ "          .envelope { border-width: 0px; border-collapse: collapse; margin:0px; padding:0px; width:700px; }\r\n"

					+ "          .envelope td { border:1px solid #126363; padding: 2px 7px 2px 7px; vertical-align:top; }\r\n"

					+ "          .fontss {font-color:white; font-family:verdana,arial,sans-serif; font-size:11px; font-weight:bold; position:relative;}\r\n"

					+ "          a:link {color: #000000; text-decoration: underline;}\r\n"

					+ "          a:active {color: #0000ff; text-decoration: underline;}\r\n"

					+ "          a:visited {color: #008000; text-decoration: underline;}\r\n"

					+ "          a:hover {color: #ff0000; text-decoration: none;}\r\n"

					+ "          .textTitleHead {color:#000; font-family:verdana,arial,sans-serif; font-size:11px; font-weight:bold;}\r\n"

					+ "          .textTitle {color:#000; font-family:verdana,arial,sans-serif; font-size:11px;}\r\n"

					+ "          .spanBold {color:black; font-family:verdana,arial,sans-serif; font-size:11px; font-weight:bold;}\r\n"

					+ "          .text8 {color:black; font-family:verdana,arial,sans-serif; font-size:11px;}\r\n"

					+ "          h4.textSekcji {font-family:verdana,arial,sans-serif; font-size:11px; font-weight:bold; text-decoration:underline; background:#CCC; margin-top:10px; margin-bottom:0px; padding:1px;}\r\n"

					+ "          .columnStyle{font-family:verdana,arial,sans-serif; font-size:11px; clolor:black; font-weight:bold}\r\n"

					+ "    #main_table {margin-left:10px; width:95%;} \r\n"

					+ "    th {background:#DDD;}\r\n"

					+ "    span.menu {font-size:16px;}\r\n"

					+ "    span.maintitle {font-size:18px; font-weight:bold;}\r\n"

					+ "    td.bline {border-top:1px solid #CCC; height:20px;}\r\n"

					+ "    ul {margin-bottom:0px; margin-top:0px;}\r\n"

					+ "    ul li, ul li span {margin-top:0px; margin bottom:0px; line-height:150%;}\r\n"

					+ "    \r\n"

					+ "\r\n"

					+ "              *{\r\n"

					+ "              font-family: Arial !important;\r\n"

					+ "              font-size: 11px;\r\n"

					+ "              cursor: default;\r\n"

					+ "              }\r\n"

					+ "\r\n"

					+ "              body {\r\n"

					+ "              font-family:Arial !important;\r\n"

					+ "              font-size: 11px;\r\n"

					+ "              cursor: default;\r\n"

					+ "              background-color:white;\r\n"

					+ "              margin: 5px;\r\n"

					+ "              color: #545559;\r\n"

					+ "              }\r\n"

					+ "\r\n"

					+ "              .title{\r\n"

					+ "              color: #0097BA;\r\n"

					+ "              padding-top:3px;\r\n"

					+ "              }\r\n"

					+ "\r\n"

					+ "              .divTableOfContents {\r\n"

					+ "              color: #0097ba;\r\n"

					+ "              margin: 15px;\r\n"

					+ "              }\r\n"

					+ "\r\n"

					+ "              td.bline {\r\n"

					+ "              border-top: 1px dotted #CCC;\r\n"

					+ "              height:0px;\r\n"

					+ "              }\r\n"

					+ "\r\n"

					+ "              .ulTableOfContents{\r\n"

					+ "              margin-bottom: 0px;\r\n"

					+ "              margin-top: 0px;\r\n"

					+ "              list-style-type: square;\r\n"

					+ "              }\r\n"

					+ "\r\n"

					+ "              .ulTableOfContents li a:link {\r\n"

					+ "              color: #0097ba;\r\n"

					+ "              text-decoration: underline;\r\n"

					+ "              font-size: 12px;\r\n"

					+ "              cursor: hand;\r\n"

					+ "              }\r\n"

					+ "\r\n"

					+ "              .ulTableOfContents li a:visited {\r\n"

					+ "              color: #0097ba;\r\n"

					+ "              }\r\n"

					+ "\r\n"

					+ "              .ulTableOfContents li a:hover {\r\n"

					+ "              color: #7FAF42;\r\n"

					+ "              }\r\n"

					+ "\r\n"

					+ "              .ulTableOfContents li a:active {\r\n"

					+ "              color: #BAD879;\r\n"

					+ "              }\r\n"

					+ "\r\n"

					+ "              .divTableOfContents h4 {\r\n"

					+ "              font-size: 14px;\r\n"

					+ "              padding: 0px;\r\n"

					+ "              margin: 5px;\r\n"

					+ "              }\r\n"

					+ "\r\n"

					+ "              .sectionDiv{\r\n"

					+ "              width:100%;\r\n"

					+ "              }\r\n"

					+ "\r\n"

					+ "              .sectionHeaderDiv {\r\n"

					+ "              background-color: #0097BA;\r\n"

					+ "              color: white;\r\n"

					+ "              padding-left: 5px;\r\n"

					+ "              width: 100%;\r\n"

					+ "              }\r\n"

					+ "\r\n"

					+ "              .sectionHeaderDiv h4 {\r\n"

					+ "              padding:3px;\r\n"

					+ "              font-size: 14px;\r\n"

					+ "              margin-bottom: 5px;\r\n"

					+ "              }\r\n"

					+ "\r\n"

					+ "              .sectionBodyDiv{\r\n"

					+ "\r\n"

					+ "              }\r\n"

					+ "\r\n"

					+ "              .backDiv{\r\n"

					+ "              text-align:right;\r\n"

					+ "              padding-top:5px;\r\n"

					+ "              padding-bottom:5px;\r\n"

					+ "              padding-right:5px;\r\n"

					+ "              color: #0097ba;\r\n"

					+ "              }\r\n"

					+ "\r\n"

					+ "              .backDiv a{\r\n"

					+ "              color: #0097ba;\r\n"

					+ "              font-size:11px;\r\n"

					+ "              font-family: Arial !important;\r\n"

					+ "              cursor:hand;\r\n"

					+ "              }\r\n"

					+ "\r\n"

					+ "              .backDiv a:visited{\r\n"

					+ "              color: #0097ba;\r\n"

					+ "              }\r\n"

					+ "\r\n"

					+ "              .backDiv a:hover{\r\n"

					+ "              color: #7FAF42;\r\n"

					+ "              }\r\n"

					+ "\r\n"

					+ "              .backDiv a:active{\r\n"

					+ "              color: #BAD879;\r\n"

					+ "              }\r\n"

					+ "\r\n"

					+ "\r\n"

					+ "              /***************************** SubSection *******************************************/\r\n"

					+ "\r\n"

					+ "              /*Subsection level 1*/\r\n"

					+ "              div.sectionDiv div.sectionBodyDiv div.subSection{\r\n"

					+ "              margin-left:15px;\r\n"

					+ "              margin-right:15px;\r\n"

					+ "              border: none;\r\n"

					+ "              margin-top:0px;\r\n"

					+ "              }\r\n"

					+ "\r\n"

					+ "              /*Subsection level 2*/\r\n"

					+ "              div.sectionDiv div.sectionBodyDiv div.subSection div.subSectionBody div.subSection {\r\n"

					+ "              margin-left:15px;\r\n"

					+ "              margin-right:15px;\r\n"

					+ "              border: none;\r\n"

					+ "              margin-top:0px;\r\n"

					+ "              }\r\n"

					+ "\r\n"

					+ "              /*Subsection level 3*/\r\n"

					+ "              div.sectionDiv div.sectionBodyDiv div.subSection  div.subSectionBody  div.subSection  div.subSectionBody  div.subSection{\r\n"

					+ "              margin-left:15px;\r\n"

					+ "              margin-right:15px;\r\n"

					+ "              border:1px dotted;\r\n"

					+ "              margin-top:2px;\r\n"

					+ "              }\r\n"

					+ "\r\n"

					+ "              /*Subsection level 4*/\r\n"

					+ "              div.sectionDiv  div.sectionBodyDiv  div.subSection  div.subSectionBody  div.subSection  div.subSectionBody  div.subSection  div.subSectionBody  div.subSection {\r\n"

					+ "              margin-left:15px;\r\n"

					+ "              margin-right:15px;\r\n"

					+ "              border: none;\r\n"

					+ "              margin-top:0px;\r\n"

					+ "              }\r\n"

					+ "\r\n"

					+ "              /*Subsection level = 5*/\r\n"

					+ "              div.subSection {\r\n"

					+ "              margin-left:15px;\r\n"

					+ "              margin-right:15px;\r\n"

					+ "              border: none;\r\n"

					+ "              margin-top:0px;\r\n"

					+ "              }\r\n"

					+ "\r\n"

					+ "              /***************************** SubSection Header *******************************************/\r\n"

					+ "\r\n"

					+ "              /*SubSection Header level 1*/\r\n"

					+ "              div.sectionDiv  div.sectionBodyDiv  div.subSection  div.subSectionHeader{\r\n"

					+ "              background-color: #85CDDB;\r\n"

					+ "              border: none;\r\n"

					+ "              color: white;\r\n"

					+ "              font-size:13px;\r\n"

					+ "              font-weight:bold;\r\n"

					+ "              padding-left: 5px;\r\n"

					+ "              }\r\n"

					+ "\r\n"

					+ "              /*SubSection Header level 2*/\r\n"

					+ "              div.sectionDiv  div.sectionBodyDiv  div.subSection  div.subSectionBody  div.subSection  div.subSectionHeader{\r\n"

					+ "              background-color: #f5f5f5;\r\n"

					+ "              border:1px solid #e6e6e6;\r\n"

					+ "              color: #545559;\r\n"

					+ "              font-size:12px;\r\n"

					+ "              font-weight:bold;\r\n"

					+ "              padding-left: 5px;\r\n"

					+ "              }\r\n"

					+ "\r\n"

					+ "              /*SubSection Header level 3*/\r\n"

					+ "              div.sectionDiv  div.sectionBodyDiv  div.subSection  div.subSectionBody  div.subSection  div.subSectionBody  div.subSection  div.subSectionHeader{\r\n"

					+ "              background-color:white;\r\n"

					+ "              border:none;\r\n"

					+ "              color: #0097ba;\r\n"

					+ "              font-size:11px;\r\n"

					+ "              font-weight:bold;\r\n"

					+ "              padding-left: 5px;\r\n"

					+ "              }\r\n"

					+ "\r\n"

					+ "              /*SubSection Header level 4*/\r\n"

					+ "              div.sectionDiv  div.sectionBodyDiv  div.subSection  div.subSectionBody  div.subSection  div.subSectionBody  div.subSection  div.subSectionBody  div.subSection  div.subSectionHeader{\r\n"

					+ "              background-color:white;\r\n"

					+ "              border:none;\r\n"

					+ "              color: #545559;\r\n"

					+ "              font-size:11px;\r\n"

					+ "              font-weight:bold;\r\n"

					+ "              padding-left: 5px;\r\n"

					+ "              }\r\n"

					+ "\r\n"

					+ "              /*SubSection Header level = 5*/\r\n"

					+ "              div.subSectionHeader {\r\n"

					+ "              background-color:white;\r\n"

					+ "              border:none;\r\n"

					+ "              color: #545559;\r\n"

					+ "              font-size:11px;\r\n"

					+ "              font-weight:bold;\r\n"

					+ "              padding-left: 5px;\r\n"

					+ "              }\r\n"

					+ "\r\n"

					+ "              /***************************** SubSection Body *******************************************/\r\n"

					+ "\r\n"

					+ "              /*SubSection Body level 1*/\r\n"

					+ "              div.sectionDiv  div.sectionBodyDiv  div.subSection  div.subSectionBody {\r\n"

					+ "              padding-left:5px;\r\n"

					+ "              padding-top:5px;\r\n"

					+ "              padding-bottom:5px;\r\n"

					+ "              padding-right:5px;\r\n"

					+ "              margin-top:0px;\r\n"

					+ "              margin-bottom:0px;\r\n"

					+ "              }\r\n"

					+ "\r\n"

					+ "              /*SubSection Body level 2*/\r\n"

					+ "              div.sectionDiv  div.sectionBodyDiv  div.subSection  div.subSectionBody  div.subSection  div.subSectionBody{\r\n"

					+ "              padding-left:5px;\r\n"

					+ "              padding-top:5px;\r\n"

					+ "              padding-bottom:5px;\r\n"

					+ "              padding-right:5px;\r\n"

					+ "              margin-top:0px;\r\n"

					+ "              margin-bottom:0px;\r\n"

					+ "              }\r\n"

					+ "\r\n"

					+ "              /*SubSection Body level 3*/\r\n"

					+ "              div.sectionDiv  div.sectionBodyDiv  div.subSection  div.subSectionBody  div.subSection  div.subSectionBody  div.subSection  div.subSectionBody{\r\n"

					+ "              padding-left:15px;\r\n"

					+ "              padding-top:0px;\r\n"

					+ "              padding-bottom:0px;\r\n"

					+ "              padding-right:0px;\r\n"

					+ "              margin-top:2px;\r\n"

					+ "              margin-bottom:2px;\r\n"

					+ "              }\r\n"

					+ "\r\n"

					+ "              /*SubSection Body level 4*/\r\n"

					+ "              div.sectionDiv  div.sectionBodyDiv  div.subSection  div.subSectionBody  div.subSection  div.subSectionBody  div.subSection  div.subSectionBody  div.subSection  div.subSectionBody{\r\n"

					+ "              padding-left:15px;\r\n"

					+ "              padding-top:0px;\r\n"

					+ "              padding-bottom:0px;\r\n"

					+ "              padding-right:0px;\r\n"

					+ "              margin-top:2px;\r\n"

					+ "              margin-bottom:2px;\r\n"

					+ "              }\r\n"

					+ "\r\n"

					+ "              /*SubSection Body level = 5*/\r\n"

					+ "              div.subSectionBody {\r\n"

					+ "              padding-left:15px;\r\n"

					+ "              padding-top:0px;\r\n"

					+ "              padding-bottom:0px;\r\n"

					+ "              padding-right:0px;\r\n"

					+ "              margin-top:2px;\r\n"

					+ "              margin-bottom:2px;\r\n"

					+ "              }\r\n"

					+ "\r\n"

					+ "              /***************************** Property *******************************************/\r\n"

					+ "\r\n"

					+ "              /*Property level 0*/\r\n"

					+ "              div.sectionDiv div.sectionBodyDiv div.property{\r\n"

					+ "              margin-left:15px;\r\n"

					+ "              margin-right:15px;\r\n"

					+ "              margin-top:2px;\r\n"

					+ "              margin-bottom: 2px;\r\n"

					+ "              padding-bottom: 2px;\r\n"

					+ "              padding-top:2px;\r\n"

					+ "              border: none;\r\n"

					+ "              }\r\n"

					+ "\r\n"

					+ "              /*Property level 1*/\r\n"

					+ "              div.sectionDiv  div.sectionBodyDiv  div.subSection  div.subSectionBody  div.property{\r\n"

					+ "              margin-left:15px;\r\n"

					+ "              margin-right:15px;\r\n"

					+ "              margin-top:2px;\r\n"

					+ "              margin-bottom: 2px;\r\n"

					+ "              padding-bottom: 2px;\r\n"

					+ "              padding-top:2px;\r\n"

					+ "              border: none;\r\n"

					+ "              }\r\n"

					+ "\r\n"

					+ "              /*Property level 2*/\r\n"

					+ "              div.sectionDiv  div.sectionBodyDiv  div.subSection  div.subSectionBody  div.subSection  div.subSectionBody  div.property {\r\n"

					+ "              margin-left:15px;\r\n"

					+ "              margin-right:15px;\r\n"

					+ "              margin-top:2px;\r\n"

					+ "              margin-bottom: 2px;\r\n"

					+ "              padding-bottom: 2px;\r\n"

					+ "              padding-top:2px;\r\n"

					+ "              border:1px dotted;\r\n"

					+ "              }\r\n"

					+ "\r\n"

					+ "              /*Property level 3*/\r\n"

					+ "              div.sectionDiv  div.sectionBodyDiv  div.subSection  div.subSectionBody  div.subSection  div.subSectionBody  div.subSection  div.subSectionBody  div.property{\r\n"

					+ "              margin-left:15px;\r\n"

					+ "              margin-right:15px;\r\n"

					+ "              margin-top:2px;\r\n"

					+ "              margin-bottom: 2px;\r\n"

					+ "              padding-bottom: 2px;\r\n"

					+ "              padding-top:2px;\r\n"

					+ "              border: none;\r\n"

					+ "              }\r\n"

					+ "\r\n"

					+ "              /*Property level = 5*/\r\n"

					+ "              div.property {\r\n"

					+ "              margin-left:15px;\r\n"

					+ "              margin-right:15px;\r\n"

					+ "              margin-top:2px;\r\n"

					+ "              margin-bottom: 2px;\r\n"

					+ "              padding-bottom: 2px;\r\n"

					+ "              padding-top:2px;\r\n"

					+ "              border: none;\r\n"

					+ "\r\n"

					+ "              }\r\n"

					+ "\r\n"

					+ "              /***************************** Property Label *******************************************/\r\n"

					+ "\r\n"

					+ "              /*SubSection Header level 1*/\r\n"

					+ "              div.sectionDiv  div.sectionBodyDiv  div.property  div.propertyLabel{\r\n"

					+ "              background-color: #85CDDB;\r\n"

					+ "              border: none;\r\n"

					+ "              padding-left: 5px;\r\n"

					+ "              font-size:13px !important;\r\n"

					+ "              font-weight:bold;\r\n"

					+ "              color: white;\r\n"

					+ "              float:none;\r\n"

					+ "              padding-right:0px !important;\r\n"

					+ "              }\r\n"

					+ "\r\n"

					+ "              div.sectionDiv  div.sectionBodyDiv  div.property  div.propertyLabel .propertyLabelSpan{\r\n"

					+ "              font-size:12px !important;\r\n"

					+ "              }\r\n"

					+ "\r\n"

					+ "              div.sectionDiv  div.sectionBodyDiv  div.property  div.propertyLabel::after{\r\n"

					+ "              content: \"\" !important;\r\n"

					+ "              }\r\n"

					+ "\r\n"

					+ "\r\n"

					+ "              /*Property Label level 1*/\r\n"

					+ "              div.sectionDiv  div.sectionBodyDiv  div.subSection  div.subSectionBody  div.property  div.propertyLabel{\r\n"

					+ "              background-color: #f5f5f5;\r\n"

					+ "              border:1px solid #e6e6e6;\r\n"

					+ "              padding-left: 5px;\r\n"

					+ "              font-size:12px !important;\r\n"

					+ "              font-weight:bold;\r\n"

					+ "              color: #545559;\r\n"

					+ "              float:none;\r\n"

					+ "              padding-right:0px !important;\r\n"

					+ "              }\r\n"

					+ "\r\n"

					+ "              div.sectionDiv  div.sectionBodyDiv  div.subSection  div.subSectionBody  div.property  div.propertyLabel .propertyLabelSpan{\r\n"

					+ "              font-size:12px !important;\r\n"

					+ "              }\r\n"

					+ "\r\n"

					+ "              div.sectionDiv  div.sectionBodyDiv  div.subSection  div.subSectionBody  div.property  div.propertyLabel::after{\r\n"

					+ "              content: \"\" !important;\r\n"

					+ "              }\r\n"

					+ "\r\n"

					+ "\r\n"

					+ "              /*Property Label level 2*/\r\n"

					+ "              div.sectionDiv  div.sectionBodyDiv  div.subSection  div.subSectionBody  div.subSection  div.subSectionBody  div.property  div.propertyLabel{\r\n"

					+ "              background-color:white;\r\n"

					+ "              border:none;\r\n"

					+ "              padding-left: 5px;\r\n"

					+ "              font-size:11px !important;\r\n"

					+ "              font-weight:bold;\r\n"

					+ "              color: #0097ba;\r\n"

					+ "              float: left;\r\n"

					+ "              padding-right:0px !important;\r\n"

					+ "              }\r\n"

					+ "\r\n"

					+ "              div.sectionDiv  div.sectionBodyDiv  div.subSection  div.subSectionBody  div.subSection  div.subSectionBody  div.property  div.propertyLabel .propertyLabelSpan{\r\n"

					+ "              font-size:11px !important;\r\n"

					+ "              }\r\n"

					+ "\r\n"

					+ "              div.sectionDiv  div.sectionBodyDiv  div.subSection  div.subSectionBody  div.subSection  div.subSectionBody  div.property  div.propertyLabel::after{\r\n"

					+ "              content: \":\" !important;\r\n"

					+ "              padding-right:5px;\r\n"

					+ "              }\r\n"

					+ "\r\n"

					+ "              /*Property Label level 3*/\r\n"

					+ "              div.sectionDiv  div.sectionBodyDiv  div.subSection  div.subSectionBody  div.subSection  div.subSectionBody  div.subSection  div.subSectionBody  div.property  div.propertyLabel{\r\n"

					+ "              background-color:white;\r\n"

					+ "              border:none;\r\n"

					+ "              padding-left: 5px;\r\n"

					+ "              font-size:11px !important;\r\n"

					+ "              font-weight:bold;\r\n"

					+ "              color: #545559;\r\n"

					+ "              float: left;\r\n"

					+ "              padding-right:0px !important;\r\n"

					+ "\r\n"

					+ "              }\r\n"

					+ "\r\n"

					+ "              div.sectionDiv  div.sectionBodyDiv  div.subSection  div.subSectionBody  div.subSection  div.subSectionBody  div.subSection  div.subSectionBody  div.property  div.propertyLabel .propertyLabelSpan{\r\n"

					+ "              font-size:11px !important;\r\n"

					+ "              }\r\n"

					+ "\r\n"

					+ "              div.sectionDiv  div.sectionBodyDiv  div.subSection  div.subSectionBody  div.subSection  div.subSectionBody  div.subSection  div.subSectionBody  div.property  div.propertyLabel::after{\r\n"

					+ "              content: \":\" !important;\r\n"

					+ "              padding-right:5px;\r\n"

					+ "              }\r\n"

					+ "\r\n"

					+ "              /*Property Label level >=4 */\r\n"

					+ "              div.propertyLabel {\r\n"

					+ "              background-color:white;\r\n"

					+ "              border:none;\r\n"

					+ "              padding-left: 5px;\r\n"

					+ "              font-size:11px !important;\r\n"

					+ "              font-weight:bold;\r\n"

					+ "              color: #545559;\r\n"

					+ "              float: left;\r\n"

					+ "              padding-right:0px !important;\r\n"

					+ "              }\r\n"

					+ "\r\n"

					+ "              div.propertyLabel .propertyLabelSpan{\r\n"

					+ "              font-size:11px !important;\r\n"

					+ "              }\r\n"

					+ "\r\n"

					+ "              div.propertyLabel::after {\r\n"

					+ "              content: \":\";\r\n"

					+ "              padding-right:5px;\r\n"

					+ "              }\r\n"

					+ "\r\n"

					+ "              /***************************** Property Value *******************************************/\r\n"

					+ "\r\n"

					+ "              /*SubSection Body level 1*/\r\n"

					+ "              div.sectionDiv  div.sectionBodyDiv  div.propertyValue {\r\n"

					+ "              padding-left:5px;\r\n"

					+ "              padding-top:5px;\r\n"

					+ "\r\n"

					+ "              }\r\n"

					+ "\r\n"

					+ "              /*Property Value level 1*/\r\n"

					+ "              div.sectionDiv  div.sectionBodyDiv  div.subSection  div.subSectionBody  div.property  div.propertyValue{\r\n"

					+ "              padding-left:5px;\r\n"

					+ "              padding-top:5px;\r\n"

					+ "              }\r\n"

					+ "\r\n"

					+ "              /*Property Value level 2*/\r\n"

					+ "              div.sectionDiv  div.sectionBodyDiv  div.subSection  div.subSectionBody  div.subSection  div.subSectionBody  div.property  div.propertyValue{\r\n"

					+ "              padding-left:5px;\r\n"

					+ "              padding-top:0px;\r\n"

					+ "              }\r\n"

					+ "\r\n"

					+ "              /*Property Value level 3*/\r\n"

					+ "              div.sectionDiv  div.sectionBodyDiv  div.subSection  div.subSectionBody  div.subSection  div.subSectionBody  div.subSection  div.subSectionBody  div.property  div.propertyValue{\r\n"

					+ "              padding-left:5px;\r\n"

					+ "              padding-top:0px;\r\n"

					+ "              }\r\n"

					+ "\r\n"

					+ "              /*Property Value level = 5*/\r\n"

					+ "              div.propertyValue {\r\n"

					+ "              padding-left:5px;\r\n"

					+ "              padding-top:0px;\r\n"

					+ "              min-width:15px;\r\n"

					+ "              min-height:14px;\r\n"

					+ "              }\r\n"

					+ "\r\n"

					+ "\r\n"

					+ "            </style>";

			htmlheader += "</head><body>";

			String htmlfooter = "</body></html>";

			printhtml.println(htmlheader);

			// start of body

			printhtml.println(
					"<table width=\"100%\"><tbody><tr id=\"topPage\"><td><div class=\"title\"><h1>Geologic Map of Scale 1:1,000,000 in "
							+ columns.get(2) + " - " + columns.get(3)
							+ "</h1></div></td></tr><tr><td class=\"bline\"></td></tr><tr><td><table width=\"100%\"><tbody><tr><td><div class=\"divTableOfContents\"><h4 xmlns=\"\">Table Of Contents</h4>\r\n"

							+ "<ul xmlns=\"\" class=\"ulTableOfContents\">\r\n"
							+ "<li><h4>General Information</h4></li>\r\n" + "<li><h4>Dataset Information</h4></li>\r\n"
							+ "<li><h4>Data Fields Description</h4></li>\r\n" + "<tr><td>"
							+ "<div id=\"generalInformation\" class=\"sectionDiv\">"
							+ "<div class=\"sectionHeaderDiv\">" + "</div>"

							//////////////////////// General Information
							//////////////////////// //////////////////////////////////////////////////////////

							+ "<div class=\"sectionBodyDiv\">"

							+ "<div xmlns=\"\" class=\"subSection\">\r\n"

							+ "<div class=\"subSectionHeader\">General Information</div>\r\n"

							+ "<div class=\"subSectionBody\"><div class=\"property\">\r\n");

			printhtml.println(GeneralDatasetInformation);

			printhtml.println("</div>\r\n"

					+ "</div>\r\n"

					+ "</div>\r\n"

					+ "</div>\r\n"

					///////////////////////////// Dataset Information
					///////////////////////////// ///////////////////////////////////////////////////

					+ "<div class=\"sectionBodyDiv\">"

					+ "<div xmlns=\"\" class=\"subSection\">\r\n"

					+ "<div class=\"subSectionHeader\">Dataset Information</div>\r\n"

					+ "<div class=\"subSectionBody\"><div class=\"property\">\r\n");

			for (int i = 0; i < DatasetInformationFirstRow.size(); i++) {
				printhtml.println("<div class=\"propertyLabel\"><span class=\"propertyLabelSpan\">"
						+ DatasetInformationFirstRow.get(i).replace("_", " ") + "</span></div>\r\n"
						+ "<div class=\"propertyValue\">" + columns.get(i) + "</div>\r\n");
			}

			printhtml.println("</div>"

					+ "</div></div>\r\n"

					//////////////////// Data Fields Description
					//////////////////// ///////////////////////////////////////

					+ "<div class=\"sectionBodyDiv\">"

					+ "<div xmlns=\"\" class=\"subSection\">\r\n"

					+ "<div class=\"subSectionHeader\">Data Fields Description</div>\r\n"

					+ "<div class=\"subSectionBody\"><div class=\"property\">\r\n"

					+ "<table>\r\n"

					+ "                           <tr>\r\n"

					+ "                             <th>Field Name</th>\r\n"

					+ "                             <th>Field Explanation</th>\r\n"

					+ "                             <th>Field Definition</th>\r\n"

					+ "                           </tr>\r\n"

					+ DataFieldsDescriptionTable

					+ "                         </table> "

					+ "</div>"

					+ "</div></div>\r\n"

					/////////////////////////////////////////////////////////////////////////////////////

					+ "</div>\r\n"

					+ "</div>\r\n"

					+ "</div>\r\n"

					+ "</div>\r\n"

					+ "</div>");

			// end of body

			printhtml.println(htmlfooter);

			printhtml.close();

			htmlfile.close();

		}

		catch (Exception e)

		{
		}

	}

}
