import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.io.PrintStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelToHtml {

	public static void main(String[] args) {
		try {
			File file = new File("C:\\Users\\92343\\Downloads\\Published_GM250K Maps Data.xlsx");
			FileInputStream fis = new FileInputStream(file);
			XSSFWorkbook wb = new XSSFWorkbook(fis);
			XSSFSheet sheet = wb.getSheetAt(0);
			Iterator<Row> itr = sheet.iterator();
			itr.next(); // Leave 1st line

			while (itr.hasNext()) {
				Row row = itr.next();
				createHtmlFile(row);
			}

			wb.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public static void createHtmlFile(Row row) {
		Iterator<Cell> cellIterator = row.cellIterator();
		List<String> columns = new ArrayList<String>();

		while (cellIterator.hasNext()) {
			Cell cell = cellIterator.next();

			if (cell.getCellType() == CellType.NUMERIC)
				columns.add(String.valueOf((int) cell.getNumericCellValue()));
			else
				columns.add(cell.getStringCellValue());
		}

		try {

			OutputStream htmlfile = new FileOutputStream(new File("c:\\Html Files\\" + columns.get(3) + ".html"));
			PrintStream printhtml = new PrintStream(htmlfile);

			String htmlheader = "<html><head>";
			htmlheader += "<title>Metadata</title> <style type=\"text/css\">\r\n" + "\r\n" + "               \r\n"
					+ "\r\n" + "* { font-family: verdana,arial,sans-serif; font-size: 11px; line-height:150%; }\r\n"
					+ "  	h1 { font-size:16px; }\r\n"
					+ "  	.header { background-color: #1C9495; color:#ffffff; font-weight: bold; }\r\n"
					+ "  	.body { background-color: #FFFFFF; font-family: verdana,arial,sans-serif; font-size: 11px; line-height:150%; }\r\n"
					+ "  	.title { text-align: left; min-width:100px; padding-left:15px; }\r\n"
					+ "  	.envelope { border-width: 0px; border-collapse: collapse; margin:0px; padding:0px; width:700px; }\r\n"
					+ "  	.envelope td { border:1px solid #126363; padding: 2px 7px 2px 7px; vertical-align:top; }\r\n"
					+ "  	.fontss {font-color:white; font-family:verdana,arial,sans-serif; font-size:11px; font-weight:bold; position:relative;}\r\n"
					+ "  	a:link {color: #000000; text-decoration: underline;}\r\n"
					+ "	a:active {color: #0000ff; text-decoration: underline;}\r\n"
					+ "	a:visited {color: #008000; text-decoration: underline;}\r\n"
					+ "	a:hover {color: #ff0000; text-decoration: none;}\r\n"
					+ "	.textTitleHead {color:#000; font-family:verdana,arial,sans-serif; font-size:11px; font-weight:bold;}\r\n"
					+ "	.textTitle {color:#000; font-family:verdana,arial,sans-serif; font-size:11px;}\r\n"
					+ "	.spanBold {color:black; font-family:verdana,arial,sans-serif; font-size:11px; font-weight:bold;}\r\n"
					+ "	.text8 {color:black; font-family:verdana,arial,sans-serif; font-size:11px;}\r\n"
					+ "	h4.textSekcji {font-family:verdana,arial,sans-serif; font-size:11px; font-weight:bold; text-decoration:underline; background:#CCC; margin-top:10px; margin-bottom:0px; padding:1px;}\r\n"
					+ "	.columnStyle{font-family:verdana,arial,sans-serif; font-size:11px; clolor:black; font-weight:bold}\r\n"
					+ "    #main_table {margin-left:10px; width:95%;} \r\n" + "    th {background:#DDD;}\r\n"
					+ "    span.menu {font-size:16px;}\r\n"
					+ "    span.maintitle {font-size:18px; font-weight:bold;}\r\n"
					+ "    td.bline {border-top:1px solid #CCC; height:20px;}\r\n"
					+ "    ul {margin-bottom:0px; margin-top:0px;}\r\n"
					+ "    ul li, ul li span {margin-top:0px; margin bottom:0px; line-height:150%;}\r\n" + "    \r\n"
					+ "\r\n" + "              *{\r\n" + "              font-family: Arial !important;\r\n"
					+ "              font-size: 11px;\r\n" + "              cursor: default;\r\n"
					+ "              }\r\n" + "\r\n" + "              body {\r\n"
					+ "              font-family:Arial !important;\r\n" + "              font-size: 11px;\r\n"
					+ "              cursor: default;\r\n" + "              background-color:white;\r\n"
					+ "              margin: 5px;\r\n" + "              color: #545559;\r\n" + "              }\r\n"
					+ "\r\n" + "              .title{\r\n" + "              color: #0097BA;\r\n"
					+ "              padding-top:3px;\r\n" + "              }\r\n" + "\r\n"
					+ "              .divTableOfContents {\r\n" + "              color: #0097ba;\r\n"
					+ "              margin: 15px;\r\n" + "              }\r\n" + "\r\n"
					+ "              td.bline {\r\n" + "              border-top: 1px dotted #CCC;\r\n"
					+ "              height:0px;\r\n" + "              }\r\n" + "\r\n"
					+ "              .ulTableOfContents{\r\n" + "              margin-bottom: 0px;\r\n"
					+ "              margin-top: 0px;\r\n" + "              list-style-type: square;\r\n"
					+ "              }\r\n" + "\r\n" + "              .ulTableOfContents li a:link {\r\n"
					+ "              color: #0097ba;\r\n" + "              text-decoration: underline;\r\n"
					+ "              font-size: 12px;\r\n" + "              cursor: hand;\r\n" + "              }\r\n"
					+ "\r\n" + "              .ulTableOfContents li a:visited {\r\n"
					+ "              color: #0097ba;\r\n" + "              }\r\n" + "\r\n"
					+ "              .ulTableOfContents li a:hover {\r\n" + "              color: #7FAF42;\r\n"
					+ "              }\r\n" + "\r\n" + "              .ulTableOfContents li a:active {\r\n"
					+ "              color: #BAD879;\r\n" + "              }\r\n" + "\r\n"
					+ "              .divTableOfContents h4 {\r\n" + "              font-size: 14px;\r\n"
					+ "              padding: 0px;\r\n" + "              margin: 5px;\r\n" + "              }\r\n"
					+ "\r\n" + "              .sectionDiv{\r\n" + "              width:100%;\r\n"
					+ "              }\r\n" + "\r\n" + "              .sectionHeaderDiv {\r\n"
					+ "              background-color: #0097BA;\r\n" + "              color: white;\r\n"
					+ "              padding-left: 5px;\r\n" + "              width: 100%;\r\n" + "              }\r\n"
					+ "\r\n" + "              .sectionHeaderDiv h4 {\r\n" + "              padding:3px;\r\n"
					+ "              font-size: 14px;\r\n" + "              margin-bottom: 5px;\r\n"
					+ "              }\r\n" + "\r\n" + "              .sectionBodyDiv{\r\n" + "\r\n"
					+ "              }\r\n" + "\r\n" + "              .backDiv{\r\n"
					+ "              text-align:right;\r\n" + "              padding-top:5px;\r\n"
					+ "              padding-bottom:5px;\r\n" + "              padding-right:5px;\r\n"
					+ "              color: #0097ba;\r\n" + "              }\r\n" + "\r\n"
					+ "              .backDiv a{\r\n" + "              color: #0097ba;\r\n"
					+ "              font-size:11px;\r\n" + "              font-family: Arial !important;\r\n"
					+ "              cursor:hand;\r\n" + "              }\r\n" + "\r\n"
					+ "              .backDiv a:visited{\r\n" + "              color: #0097ba;\r\n"
					+ "              }\r\n" + "\r\n" + "              .backDiv a:hover{\r\n"
					+ "              color: #7FAF42;\r\n" + "              }\r\n" + "\r\n"
					+ "              .backDiv a:active{\r\n" + "              color: #BAD879;\r\n"
					+ "              }\r\n" + "\r\n" + "\r\n"
					+ "              /***************************** SubSection *******************************************/\r\n"
					+ "\r\n" + "              /*Subsection level 1*/\r\n"
					+ "              div.sectionDiv div.sectionBodyDiv div.subSection{\r\n"
					+ "              margin-left:15px;\r\n" + "              margin-right:15px;\r\n"
					+ "              border: none;\r\n" + "              margin-top:0px;\r\n" + "              }\r\n"
					+ "\r\n" + "              /*Subsection level 2*/\r\n"
					+ "              div.sectionDiv div.sectionBodyDiv div.subSection div.subSectionBody div.subSection {\r\n"
					+ "              margin-left:15px;\r\n" + "              margin-right:15px;\r\n"
					+ "              border: none;\r\n" + "              margin-top:0px;\r\n" + "              }\r\n"
					+ "\r\n" + "              /*Subsection level 3*/\r\n"
					+ "              div.sectionDiv div.sectionBodyDiv div.subSection  div.subSectionBody  div.subSection  div.subSectionBody  div.subSection{\r\n"
					+ "              margin-left:15px;\r\n" + "              margin-right:15px;\r\n"
					+ "              border:1px dotted;\r\n" + "              margin-top:2px;\r\n"
					+ "              }\r\n" + "\r\n" + "              /*Subsection level 4*/\r\n"
					+ "              div.sectionDiv  div.sectionBodyDiv  div.subSection  div.subSectionBody  div.subSection  div.subSectionBody  div.subSection  div.subSectionBody  div.subSection {\r\n"
					+ "              margin-left:15px;\r\n" + "              margin-right:15px;\r\n"
					+ "              border: none;\r\n" + "              margin-top:0px;\r\n" + "              }\r\n"
					+ "\r\n" + "              /*Subsection level = 5*/\r\n" + "              div.subSection {\r\n"
					+ "              margin-left:15px;\r\n" + "              margin-right:15px;\r\n"
					+ "              border: none;\r\n" + "              margin-top:0px;\r\n" + "              }\r\n"
					+ "\r\n"
					+ "              /***************************** SubSection Header *******************************************/\r\n"
					+ "\r\n" + "              /*SubSection Header level 1*/\r\n"
					+ "              div.sectionDiv  div.sectionBodyDiv  div.subSection  div.subSectionHeader{\r\n"
					+ "              background-color: #85CDDB;\r\n" + "              border: none;\r\n"
					+ "              color: white;\r\n" + "              font-size:13px;\r\n"
					+ "              font-weight:bold;\r\n" + "              padding-left: 5px;\r\n"
					+ "              }\r\n" + "\r\n" + "              /*SubSection Header level 2*/\r\n"
					+ "              div.sectionDiv  div.sectionBodyDiv  div.subSection  div.subSectionBody  div.subSection  div.subSectionHeader{\r\n"
					+ "              background-color: #f5f5f5;\r\n" + "              border:1px solid #e6e6e6;\r\n"
					+ "              color: #545559;\r\n" + "              font-size:12px;\r\n"
					+ "              font-weight:bold;\r\n" + "              padding-left: 5px;\r\n"
					+ "              }\r\n" + "\r\n" + "              /*SubSection Header level 3*/\r\n"
					+ "              div.sectionDiv  div.sectionBodyDiv  div.subSection  div.subSectionBody  div.subSection  div.subSectionBody  div.subSection  div.subSectionHeader{\r\n"
					+ "              background-color:white;\r\n" + "              border:none;\r\n"
					+ "              color: #0097ba;\r\n" + "              font-size:11px;\r\n"
					+ "              font-weight:bold;\r\n" + "              padding-left: 5px;\r\n"
					+ "              }\r\n" + "\r\n" + "              /*SubSection Header level 4*/\r\n"
					+ "              div.sectionDiv  div.sectionBodyDiv  div.subSection  div.subSectionBody  div.subSection  div.subSectionBody  div.subSection  div.subSectionBody  div.subSection  div.subSectionHeader{\r\n"
					+ "              background-color:white;\r\n" + "              border:none;\r\n"
					+ "              color: #545559;\r\n" + "              font-size:11px;\r\n"
					+ "              font-weight:bold;\r\n" + "              padding-left: 5px;\r\n"
					+ "              }\r\n" + "\r\n" + "              /*SubSection Header level = 5*/\r\n"
					+ "              div.subSectionHeader {\r\n" + "              background-color:white;\r\n"
					+ "              border:none;\r\n" + "              color: #545559;\r\n"
					+ "              font-size:11px;\r\n" + "              font-weight:bold;\r\n"
					+ "              padding-left: 5px;\r\n" + "              }\r\n" + "\r\n"
					+ "              /***************************** SubSection Body *******************************************/\r\n"
					+ "\r\n" + "              /*SubSection Body level 1*/\r\n"
					+ "              div.sectionDiv  div.sectionBodyDiv  div.subSection  div.subSectionBody {\r\n"
					+ "              padding-left:5px;\r\n" + "              padding-top:5px;\r\n"
					+ "              padding-bottom:5px;\r\n" + "              padding-right:5px;\r\n"
					+ "              margin-top:0px;\r\n" + "              margin-bottom:0px;\r\n"
					+ "              }\r\n" + "\r\n" + "              /*SubSection Body level 2*/\r\n"
					+ "              div.sectionDiv  div.sectionBodyDiv  div.subSection  div.subSectionBody  div.subSection  div.subSectionBody{\r\n"
					+ "              padding-left:5px;\r\n" + "              padding-top:5px;\r\n"
					+ "              padding-bottom:5px;\r\n" + "              padding-right:5px;\r\n"
					+ "              margin-top:0px;\r\n" + "              margin-bottom:0px;\r\n"
					+ "              }\r\n" + "\r\n" + "              /*SubSection Body level 3*/\r\n"
					+ "              div.sectionDiv  div.sectionBodyDiv  div.subSection  div.subSectionBody  div.subSection  div.subSectionBody  div.subSection  div.subSectionBody{\r\n"
					+ "              padding-left:15px;\r\n" + "              padding-top:0px;\r\n"
					+ "              padding-bottom:0px;\r\n" + "              padding-right:0px;\r\n"
					+ "              margin-top:2px;\r\n" + "              margin-bottom:2px;\r\n"
					+ "              }\r\n" + "\r\n" + "              /*SubSection Body level 4*/\r\n"
					+ "              div.sectionDiv  div.sectionBodyDiv  div.subSection  div.subSectionBody  div.subSection  div.subSectionBody  div.subSection  div.subSectionBody  div.subSection  div.subSectionBody{\r\n"
					+ "              padding-left:15px;\r\n" + "              padding-top:0px;\r\n"
					+ "              padding-bottom:0px;\r\n" + "              padding-right:0px;\r\n"
					+ "              margin-top:2px;\r\n" + "              margin-bottom:2px;\r\n"
					+ "              }\r\n" + "\r\n" + "              /*SubSection Body level = 5*/\r\n"
					+ "              div.subSectionBody {\r\n" + "              padding-left:15px;\r\n"
					+ "              padding-top:0px;\r\n" + "              padding-bottom:0px;\r\n"
					+ "              padding-right:0px;\r\n" + "              margin-top:2px;\r\n"
					+ "              margin-bottom:2px;\r\n" + "              }\r\n" + "\r\n"
					+ "              /***************************** Property *******************************************/\r\n"
					+ "\r\n" + "              /*Property level 0*/\r\n"
					+ "              div.sectionDiv div.sectionBodyDiv div.property{\r\n"
					+ "              margin-left:15px;\r\n" + "              margin-right:15px;\r\n"
					+ "              margin-top:2px;\r\n" + "              margin-bottom: 2px;\r\n"
					+ "              padding-bottom: 2px;\r\n" + "              padding-top:2px;\r\n"
					+ "              border: none;\r\n" + "              }\r\n" + "\r\n"
					+ "              /*Property level 1*/\r\n"
					+ "              div.sectionDiv  div.sectionBodyDiv  div.subSection  div.subSectionBody  div.property{\r\n"
					+ "              margin-left:15px;\r\n" + "              margin-right:15px;\r\n"
					+ "              margin-top:2px;\r\n" + "              margin-bottom: 2px;\r\n"
					+ "              padding-bottom: 2px;\r\n" + "              padding-top:2px;\r\n"
					+ "              border: none;\r\n" + "              }\r\n" + "\r\n"
					+ "              /*Property level 2*/\r\n"
					+ "              div.sectionDiv  div.sectionBodyDiv  div.subSection  div.subSectionBody  div.subSection  div.subSectionBody  div.property {\r\n"
					+ "              margin-left:15px;\r\n" + "              margin-right:15px;\r\n"
					+ "              margin-top:2px;\r\n" + "              margin-bottom: 2px;\r\n"
					+ "              padding-bottom: 2px;\r\n" + "              padding-top:2px;\r\n"
					+ "              border:1px dotted;\r\n" + "              }\r\n" + "\r\n"
					+ "              /*Property level 3*/\r\n"
					+ "              div.sectionDiv  div.sectionBodyDiv  div.subSection  div.subSectionBody  div.subSection  div.subSectionBody  div.subSection  div.subSectionBody  div.property{\r\n"
					+ "              margin-left:15px;\r\n" + "              margin-right:15px;\r\n"
					+ "              margin-top:2px;\r\n" + "              margin-bottom: 2px;\r\n"
					+ "              padding-bottom: 2px;\r\n" + "              padding-top:2px;\r\n"
					+ "              border: none;\r\n" + "              }\r\n" + "\r\n"
					+ "              /*Property level = 5*/\r\n" + "              div.property {\r\n"
					+ "              margin-left:15px;\r\n" + "              margin-right:15px;\r\n"
					+ "              margin-top:2px;\r\n" + "              margin-bottom: 2px;\r\n"
					+ "              padding-bottom: 2px;\r\n" + "              padding-top:2px;\r\n"
					+ "              border: none;\r\n" + "\r\n" + "              }\r\n" + "\r\n"
					+ "              /***************************** Property Label *******************************************/\r\n"
					+ "\r\n" + "              /*SubSection Header level 1*/\r\n"
					+ "              div.sectionDiv  div.sectionBodyDiv  div.property  div.propertyLabel{\r\n"
					+ "              background-color: #85CDDB;\r\n" + "              border: none;\r\n"
					+ "              padding-left: 5px;\r\n" + "              font-size:13px !important;\r\n"
					+ "              font-weight:bold;\r\n" + "              color: white;\r\n"
					+ "              float:none;\r\n" + "              padding-right:0px !important;\r\n"
					+ "              }\r\n" + "\r\n"
					+ "              div.sectionDiv  div.sectionBodyDiv  div.property  div.propertyLabel .propertyLabelSpan{\r\n"
					+ "              font-size:12px !important;\r\n" + "              }\r\n" + "\r\n"
					+ "              div.sectionDiv  div.sectionBodyDiv  div.property  div.propertyLabel::after{\r\n"
					+ "              content: \"\" !important;\r\n" + "              }\r\n" + "\r\n" + "\r\n"
					+ "              /*Property Label level 1*/\r\n"
					+ "              div.sectionDiv  div.sectionBodyDiv  div.subSection  div.subSectionBody  div.property  div.propertyLabel{\r\n"
					+ "              background-color: #f5f5f5;\r\n" + "              border:1px solid #e6e6e6;\r\n"
					+ "              padding-left: 5px;\r\n" + "              font-size:12px !important;\r\n"
					+ "              font-weight:bold;\r\n" + "              color: #545559;\r\n"
					+ "              float:none;\r\n" + "              padding-right:0px !important;\r\n"
					+ "              }\r\n" + "\r\n"
					+ "              div.sectionDiv  div.sectionBodyDiv  div.subSection  div.subSectionBody  div.property  div.propertyLabel .propertyLabelSpan{\r\n"
					+ "              font-size:12px !important;\r\n" + "              }\r\n" + "\r\n"
					+ "              div.sectionDiv  div.sectionBodyDiv  div.subSection  div.subSectionBody  div.property  div.propertyLabel::after{\r\n"
					+ "              content: \"\" !important;\r\n" + "              }\r\n" + "\r\n" + "\r\n"
					+ "              /*Property Label level 2*/\r\n"
					+ "              div.sectionDiv  div.sectionBodyDiv  div.subSection  div.subSectionBody  div.subSection  div.subSectionBody  div.property  div.propertyLabel{\r\n"
					+ "              background-color:white;\r\n" + "              border:none;\r\n"
					+ "              padding-left: 5px;\r\n" + "              font-size:11px !important;\r\n"
					+ "              font-weight:bold;\r\n" + "              color: #0097ba;\r\n"
					+ "              float: left;\r\n" + "              padding-right:0px !important;\r\n"
					+ "              }\r\n" + "\r\n"
					+ "              div.sectionDiv  div.sectionBodyDiv  div.subSection  div.subSectionBody  div.subSection  div.subSectionBody  div.property  div.propertyLabel .propertyLabelSpan{\r\n"
					+ "              font-size:11px !important;\r\n" + "              }\r\n" + "\r\n"
					+ "              div.sectionDiv  div.sectionBodyDiv  div.subSection  div.subSectionBody  div.subSection  div.subSectionBody  div.property  div.propertyLabel::after{\r\n"
					+ "              content: \":\" !important;\r\n" + "              padding-right:5px;\r\n"
					+ "              }\r\n" + "\r\n" + "              /*Property Label level 3*/\r\n"
					+ "              div.sectionDiv  div.sectionBodyDiv  div.subSection  div.subSectionBody  div.subSection  div.subSectionBody  div.subSection  div.subSectionBody  div.property  div.propertyLabel{\r\n"
					+ "              background-color:white;\r\n" + "              border:none;\r\n"
					+ "              padding-left: 5px;\r\n" + "              font-size:11px !important;\r\n"
					+ "              font-weight:bold;\r\n" + "              color: #545559;\r\n"
					+ "              float: left;\r\n" + "              padding-right:0px !important;\r\n" + "\r\n"
					+ "              }\r\n" + "\r\n"
					+ "              div.sectionDiv  div.sectionBodyDiv  div.subSection  div.subSectionBody  div.subSection  div.subSectionBody  div.subSection  div.subSectionBody  div.property  div.propertyLabel .propertyLabelSpan{\r\n"
					+ "              font-size:11px !important;\r\n" + "              }\r\n" + "\r\n"
					+ "              div.sectionDiv  div.sectionBodyDiv  div.subSection  div.subSectionBody  div.subSection  div.subSectionBody  div.subSection  div.subSectionBody  div.property  div.propertyLabel::after{\r\n"
					+ "              content: \":\" !important;\r\n" + "              padding-right:5px;\r\n"
					+ "              }\r\n" + "\r\n" + "              /*Property Label level >=4 */\r\n"
					+ "              div.propertyLabel {\r\n" + "              background-color:white;\r\n"
					+ "              border:none;\r\n" + "              padding-left: 5px;\r\n"
					+ "              font-size:11px !important;\r\n" + "              font-weight:bold;\r\n"
					+ "              color: #545559;\r\n" + "              float: left;\r\n"
					+ "              padding-right:0px !important;\r\n" + "              }\r\n" + "\r\n"
					+ "              div.propertyLabel .propertyLabelSpan{\r\n"
					+ "              font-size:11px !important;\r\n" + "              }\r\n" + "\r\n"
					+ "              div.propertyLabel::after {\r\n" + "              content: \":\";\r\n"
					+ "              padding-right:5px;\r\n" + "              }\r\n" + "\r\n"
					+ "              /***************************** Property Value *******************************************/\r\n"
					+ "\r\n" + "              /*SubSection Body level 1*/\r\n"
					+ "              div.sectionDiv  div.sectionBodyDiv  div.propertyValue {\r\n"
					+ "              padding-left:5px;\r\n" + "              padding-top:5px;\r\n" + "\r\n"
					+ "              }\r\n" + "\r\n" + "              /*Property Value level 1*/\r\n"
					+ "              div.sectionDiv  div.sectionBodyDiv  div.subSection  div.subSectionBody  div.property  div.propertyValue{\r\n"
					+ "              padding-left:5px;\r\n" + "              padding-top:5px;\r\n"
					+ "              }\r\n" + "\r\n" + "              /*Property Value level 2*/\r\n"
					+ "              div.sectionDiv  div.sectionBodyDiv  div.subSection  div.subSectionBody  div.subSection  div.subSectionBody  div.property  div.propertyValue{\r\n"
					+ "              padding-left:5px;\r\n" + "              padding-top:0px;\r\n"
					+ "              }\r\n" + "\r\n" + "              /*Property Value level 3*/\r\n"
					+ "              div.sectionDiv  div.sectionBodyDiv  div.subSection  div.subSectionBody  div.subSection  div.subSectionBody  div.subSection  div.subSectionBody  div.property  div.propertyValue{\r\n"
					+ "              padding-left:5px;\r\n" + "              padding-top:0px;\r\n"
					+ "              }\r\n" + "\r\n" + "              /*Property Value level = 5*/\r\n"
					+ "              div.propertyValue {\r\n" + "              padding-left:5px;\r\n"
					+ "              padding-top:0px;\r\n" + "              min-width:15px;\r\n"
					+ "              min-height:14px;\r\n" + "              }\r\n" + "\r\n" + "\r\n"
					+ "            </style>";
			htmlheader += "</head><body>";
			String htmlfooter = "</body></html>";

			printhtml.println(htmlheader);

			// start of body

			/*
			 * columns.get(0) Object ID columns.get(1) Geo Name English columns.get(2) Geo
			 * Name Arabic columns.get(3) Pub No columns.get(4) Affiliate columns.get(5)
			 * Publisher columns.get(6) Compile Period columns.get(7) Publish Date
			 * columns.get(8) Title columns.get(9) Summary columns.get(10) Index
			 * columns.get(11) Note columns.get(12)Authors columns.get(13) Area Type
			 * columns.get(14) Geology columns.get(15) Author 1 columns.get(16) Author 2
			 * columns.get(17) Author 3 columns.get(18) Min Latitude columns.get(19) Min
			 * Longitude columns.get(20) Max Latitude columns.get(21) Max Longitude
			 */

			printhtml.println("<table width=\"100%\"><tbody><tr id=\"topPage\"><td><div class=\"title\"><h1>"
					+ columns.get(8)
					+ "</h1></div></td></tr><tr><td class=\"bline\"></td></tr><tr><td><table width=\"100%\"><tbody><tr><td><div class=\"divTableOfContents\"><h4 xmlns=\"\">Table Of Contents</h4>\r\n"
					+ "<ul xmlns=\"\" class=\"ulTableOfContents\">\r\n"
					+ "<li><a href=\"\">Dataset Information</a></li>\r\n"
					+ "<li><a href=\"\">Compilation Period</a></li>\r\n"
					+ "<li><a href=\"\">Contact Information</a></li>\r\n" + "<li><a href=\"\">Access</a></li>\r\n"
					+ "</ul></div></td><td class=\"overviewImg\"><div id=\"thumbnailSpan\"><img src=\"https://forum.huawei.com/enterprise/en/data/attachment/forum/201909/27/100129gvscd0mzd7ul9ukr.jpg?GIS.JPG\" alt=\"GIS Image\" width=\"200\" height=\"200\" onerror=\"this.src=&#39;img/no_thumbnail.png&#39;\"></div></td></tr></tbody></table></td></tr>"

					+ "<tr><td><div id=\"generalInformation\" class=\"sectionDiv\"><div class=\"sectionHeaderDiv\"><h4>Content</h4></div>"
					+ "<div class=\"sectionBodyDiv\">" + "<div xmlns=\"\" class=\"subSection\">\r\n"

					+ "<div class=\"subSectionHeader\">Dataset Information</div>\r\n"
					+ "<div class=\"subSectionBody\"><div class=\"property\">\r\n"

					+ "<div class=\"propertyLabel\"><span class=\"propertyLabelSpan\">Publish No</span></div>\r\n"
					+ "<div class=\"propertyValue\">" + columns.get(3) + "</div>\r\n"

					+ "<div class=\"propertyLabel\"><span class=\"propertyLabelSpan\">Geo Name English</span></div>\r\n"
					+ "<div class=\"propertyValue\">" + columns.get(1) + "</div>\r\n"

					+ "<div class=\"propertyLabel\"><span class=\"propertyLabelSpan\">Geo Name Arabic</span></div>\r\n"
					+ "<div class=\"propertyValue\">" + columns.get(2) + "</div>\r\n"

					+ "<div class=\"propertyLabel\"><span class=\"propertyLabelSpan\">Affiliate</span></div>\r\n"
					+ "<div class=\"propertyValue\">" + columns.get(4) + "</div>\r\n"

					+ "<div class=\"propertyLabel\"><span class=\"propertyLabelSpan\">Publisher</span></div>\r\n"
					+ "<div class=\"propertyValue\">" + columns.get(5) + "</div>\r\n"

//					+ "<div class=\"propertyLabel\"><span class=\"propertyLabelSpan\">Compile Period</span></div>\r\n"
//					+ "<div class=\"propertyValue\">"+columns.get(6)+"</div>\r\n"
//					
//					+ "<div class=\"propertyLabel\"><span class=\"propertyLabelSpan\">Publish Date</span></div>\r\n"
//					+ "<div class=\"propertyValue\">"+columns.get(7)+"</div>\r\n"

					+ "<div class=\"propertyLabel\"><span class=\"propertyLabelSpan\">Title</span></div>\r\n"
					+ "<div class=\"propertyValue\">" + columns.get(8) + "</div>\r\n"

					+ "<div class=\"propertyLabel\"><span class=\"propertyLabelSpan\">Summary</span></div>\r\n"
					+ "<div class=\"propertyValue\">" + columns.get(9) + "</div>\r\n"

					+ "<div class=\"propertyLabel\"><span class=\"propertyLabelSpan\">Index</span></div>\r\n"
					+ "<div class=\"propertyValue\">" + columns.get(10) + "</div>\r\n"

					+ "<div class=\"propertyLabel\"><span class=\"propertyLabelSpan\">Note</span></div>\r\n"
					+ "<div class=\"propertyValue\">" + columns.get(11) + "</div>\r\n"

					+ "<div class=\"propertyLabel\"><span class=\"propertyLabelSpan\">Authors</span></div>\r\n"
					+ "<div class=\"propertyValue\">" + columns.get(12) + "</div>\r\n"

					+ "<div class=\"propertyLabel\"><span class=\"propertyLabelSpan\">Area Type</span></div>\r\n"
					+ "<div class=\"propertyValue\">" + columns.get(13) + "</div>\r\n"

					+ "<div class=\"propertyLabel\"><span class=\"propertyLabelSpan\">Geology</span></div>\r\n"
					+ "<div class=\"propertyValue\">" + columns.get(14) + "</div>\r\n"

					+ "<div class=\"propertyLabel\"><span class=\"propertyLabelSpan\">Authors 1</span></div>\r\n"
					+ "<div class=\"propertyValue\">" + columns.get(15) + "</div>\r\n"

					+ "<div class=\"propertyLabel\"><span class=\"propertyLabelSpan\">Authors 2</span></div>\r\n"
					+ "<div class=\"propertyValue\">" + columns.get(16) + "</div>\r\n"

					+ "<div class=\"propertyLabel\"><span class=\"propertyLabelSpan\">Authors 3</span></div>\r\n"
					+ "<div class=\"propertyValue\">" + columns.get(17) + "</div>\r\n"

					+ "<div class=\"propertyLabel\"><span class=\"propertyLabelSpan\">Min Latitude</span></div>\r\n"
					+ "<div class=\"propertyValue\">" + columns.get(18) + "</div>\r\n"

					+ "<div class=\"propertyLabel\"><span class=\"propertyLabelSpan\">Min Longitude</span></div>\r\n"
					+ "<div class=\"propertyValue\">" + columns.get(19) + "</div>\r\n"

					+ "<div class=\"propertyLabel\"><span class=\"propertyLabelSpan\">Max Latitude</span></div>\r\n"
					+ "<div class=\"propertyValue\">" + columns.get(20) + "</div>\r\n"

					+ "<div class=\"propertyLabel\"><span class=\"propertyLabelSpan\">Max Longitude</span></div>\r\n"
					+ "<div class=\"propertyValue\">" + columns.get(21) + "</div>\r\n"

					+ "</div></div>\r\n"

					+ "<div xmlns=\"\" class=\"subSection\">\r\n"
					+ "<div class=\"subSectionHeader\">Compilation Period</div>\r\n"
					+ "<div class=\"subSectionBody\">\r\n" + "<div class=\"property\">\r\n"
					+ "<div class=\"propertyLabel\"><span class=\"propertyLabelSpan\">Date</span></div>\r\n"
					+ "<div class=\"propertyValue\">" + columns.get(6) + "</div>\r\n" + "</div>\r\n"
					+ "<div class=\"property\">\r\n"
					+ "<div class=\"propertyLabel\"><span class=\"propertyLabelSpan\">Date Type</span></div>\r\n"
					+ "<div class=\"propertyValue\">Creation</div>\r\n" + "</div>\r\n" + "</div>\r\n" + "</div>\r\n"
					+ "<div id=\"contact\" class=\"sectionDiv\"><div class=\"sectionHeaderDiv\"><h4>Contact Information</h4></div><div class=\"sectionBodyDiv\"><div xmlns=\"\" class=\"subSection\">\r\n"
					+ "<div class=\"subSectionHeader\">Contact</div>\r\n" + "<div class=\"subSectionBody\">\r\n"
					+ "<div class=\"property\">\r\n"
					+ "<div class=\"propertyLabel\"><span class=\"propertyLabelSpan\">Role</span></div>\r\n"
					+ "<div class=\"propertyValue\">Point Of Contact</div>\r\n" + "</div>\r\n"
					+ "<div class=\"subSection\">\r\n" + "<div class=\"subSectionHeader\">Organisation Party</div>\r\n"
					+ "<div class=\"subSectionBody\">\r\n" + "<div class=\"property\">\r\n"
					+ "<div class=\"propertyLabel\"><span class=\"propertyLabelSpan\">Name</span></div>\r\n"
					+ "<div class=\"propertyValue\">Hexagon Geospatial</div>\r\n" + "</div>\r\n"
					+ "<div class=\"subSection\">\r\n" + "<div class=\"subSectionHeader\">Individual</div>\r\n"
					+ "<div class=\"subSectionBody\">\r\n" + "<div class=\"property\">\r\n"
					+ "<div class=\"propertyLabel\"><span class=\"propertyLabelSpan\">Name</span></div>\r\n"
					+ "<div class=\"propertyValue\">Hexagon Geospatial</div>\r\n" + "</div>\r\n"
					+ "<div class=\"subSection\">\r\n" + "<div class=\"subSectionHeader\">Contact Info</div>\r\n"
					+ "<div class=\"subSectionBody\">\r\n" + "<div class=\"subSection\">\r\n"
					+ "<div class=\"subSectionHeader\">Phone</div>\r\n" + "<div class=\"subSectionBody\">\r\n"
					+ "<div class=\"property\">\r\n"
					+ "<div class=\"propertyLabel\"><span class=\"propertyLabelSpan\">Number</span></div>\r\n"
					+ "<div class=\"propertyValue\">+966 12 619 5000</div>\r\n" + "</div>\r\n"
					+ "<div class=\"property\">\r\n"
					+ "<div class=\"propertyLabel\"><span class=\"propertyLabelSpan\">Number Type</span></div>\r\n"
					+ "<div class=\"propertyValue\">Landline Voice</div>\r\n" + "</div>\r\n" + "</div>\r\n"
					+ "</div>\r\n" + "<div class=\"subSection\">\r\n"
					+ "<div class=\"subSectionHeader\">Phone</div>\r\n" + "<div class=\"subSectionBody\">\r\n"
					+ "<div class=\"property\">\r\n"
					+ "<div class=\"propertyLabel\"><span class=\"propertyLabelSpan\">Number</span></div>\r\n"
					+ "<div class=\"propertyValue\">+966 12 619 6000</div>\r\n" + "</div>\r\n"
					+ "<div class=\"property\">\r\n"
					+ "<div class=\"propertyLabel\"><span class=\"propertyLabelSpan\">Number Type</span></div>\r\n"
					+ "<div class=\"propertyValue\">Fax</div>\r\n" + "</div>\r\n" + "</div>\r\n" + "</div>\r\n"
					+ "<div class=\"subSection\">\r\n" + "<div class=\"subSectionHeader\">Address</div>\r\n"
					+ "<div class=\"subSectionBody\">\r\n" + "<div class=\"property\">\r\n"
					+ "<div class=\"propertyLabel\"><span class=\"propertyLabelSpan\">Delivery Point</span></div>\r\n"
					+ "<div class=\"propertyValue\">Ahmed bin Mohammed Al Ashab, Al Waha, Jeddah</div>\r\n"
					+ "</div>\r\n" + "<div class=\"property\">\r\n"
					+ "<div class=\"propertyLabel\"><span class=\"propertyLabelSpan\">City</span></div>\r\n"
					+ "<div class=\"propertyValue\">Jeddah</div>\r\n" + "</div>\r\n" + "<div class=\"property\">\r\n"
					+ "<div class=\"propertyLabel\"><span class=\"propertyLabelSpan\">Postal Code</span></div>\r\n"
					+ "<div class=\"propertyValue\">54141 Jeddah 21514</div>\r\n" + "</div>\r\n"
					+ "<div class=\"property\">\r\n"
					+ "<div class=\"propertyLabel\"><span class=\"propertyLabelSpan\">Country</span></div>\r\n"
					+ "<div class=\"propertyValue\">Saudi Arabia</div>\r\n" + "</div>\r\n"
					+ "<div class=\"property\">\r\n"
					+ "<div class=\"propertyLabel\"><span class=\"propertyLabelSpan\">Electronic Mail Address</span></div>\r\n"
					+ "<div class=\"propertyValue\">ngdwebmaster@sgs.org.sa</div>\r\n" + "</div>\r\n" + "</div>\r\n"
					+ "</div>\r\n" + "</div>\r\n" + "</div>\r\n" + "<div class=\"property\">\r\n"
					+ "<div class=\"propertyLabel\"><span class=\"propertyLabelSpan\">Position Name</span></div>\r\n"
					+ "<div class=\"propertyValue\">National Geological Database Webmaster</div>\r\n" + "</div>\r\n"
					+ "</div>\r\n" + "</div>\r\n" + "</div>\r\n" + "</div>\r\n" + "</div>\r\n" + "</div>"

					+ "<div id=\"contact\" class=\"sectionDiv\"><div class=\"sectionHeaderDiv\"><h4>Access</h4></div><div class=\"sectionBodyDiv\"><div xmlns=\"\" class=\"subSection\">\r\n"
					+ "<div class=\"subSectionHeader\">Restrictions</div>\r\n" + "<div class=\"subSectionBody\">\r\n"
					+ "<div class=\"property\">\r\n"
					+ "<div class=\"propertyLabel\"><span class=\"propertyLabelSpan\">Use Limitations</span></div>\r\n"
					+ "<div class=\"propertyValue\">Restricted License</div>\r\n" + "</div>\r\n"
					+ "<div class=\"subSectionBody\">\r\n" + "<div class=\"property\">\r\n"
					+ "<div class=\"propertyLabel\"><span class=\"propertyLabelSpan\">Access Constraints</span></div>\r\n"
					+ "<div class=\"propertyValue\">No Restictions</div>\r\n" + "</div>\r\n"
					+ "<div class=\"subSectionBody\">\r\n" + "<div class=\"property\">\r\n"
					+ "<div class=\"propertyLabel\"><span class=\"propertyLabelSpan\">Other Constraints</span></div>\r\n"
					+ "<div class=\"propertyValue\">No Other Constraints</div>\r\n" + "</div>\r\n"

					+ "</div>\r\n" + "</div>\r\n" + "</div>\r\n" + "</div>\r\n" + "</div>\r\n" + "</div>"
//            		+ "<div xmlns=\"\" class=\"backDiv\">"
//            		+ "<a href=\"https://ngp.sgs.org.sa/XsltApplicator.WebClient.ashx?action=ErdasHtml&amp;link=http%3A%2F%2Fngdp.sgs.org.sa%2F%2Ferdas-apollo%2Fcatalog%2Fcontent%2Fitems%2F2c949881752363fd017523a1552f21a4%2Fattachment%2Fdefault&amp;xsltUrl=XsltGenerator.WebClient.ashx%3Faction%3DErdasHtml%26title%3DGeology%20of%20the%20Nuqrah%20Quadrangle%2C%20Sheet%2025E.%26apolloUrl%3Dhttp%3A%2F%2Fngdp.sgs.org.sa%2F%26id%3D2c949881752363fd017523a1552f21a4%26mapServiceId%3D492e7233-11aa-409b-aef5-9ffb48b03b34#topPage\">Back To Table Of Contents</a></div></div></div></td></tr>"	
					+ "");

			// end of body

			printhtml.println(htmlfooter);

			printhtml.close();
			htmlfile.close();
		}

		catch (Exception e) {
		}

	}

}
