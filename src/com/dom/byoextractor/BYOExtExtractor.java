package com.dom.byoextractor;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Connection;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

public class BYOExtExtractor {
	int rowCount = 0;
	Element defaults,paints,wheels,optPkgs,facInstOpt,acc;
	Element defaultsTblContent, paintsTblcontent, wheelsTblContent, optPkgsTblcontent,optFactInstTblContent,accTblContent;
	private Row rowHeader;
	private Cell cell = null;
	List<String> pkgNames = new ArrayList<String>();

	@SuppressWarnings("deprecation")
	public void tblIterator(XSSFWorkbook wb, XSSFSheet sheet,String title, Element tbl){
		//Set Header Style
		XSSFCellStyle headerStyle = wb.createCellStyle();
		headerStyle.setFillForegroundColor(new XSSFColor(new java.awt.Color(22, 54, 92)));
		headerStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
		headerStyle.setAlignment(CellStyle.ALIGN_CENTER);

		//Set Sub-Header Style
		XSSFCellStyle subHeaderStyle = wb.createCellStyle();
		subHeaderStyle.setFillForegroundColor(new XSSFColor(new java.awt.Color(197, 217, 241)));
		subHeaderStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);

		//Set Header font color
		Font headerFont = wb.createFont();
		headerFont.setFontName("Calibri");
		headerFont.setBoldweight(Font.BOLDWEIGHT_NORMAL);
		headerFont.setColor(IndexedColors.WHITE.getIndex());
		headerStyle.setFont(headerFont);

		rowHeader = sheet.createRow(rowCount);
		cell = rowHeader.createCell(0);
		cell.setCellValue(title.toLowerCase());
		cell.setCellStyle(headerStyle);
		sheet.addMergedRegion(new CellRangeAddress(rowCount, rowCount, 0, 4));
		rowCount++;
		for (Element row : tbl.select("tr")) {
			rowHeader = sheet.createRow(rowCount);
			// loop through all
			Elements ths = row.select("th");
			int count = 0;
			for (Element thContent : ths) {
				// set header style
				cell = rowHeader.createCell(count);
				cell.setCellValue(thContent.text());
				cell.setCellStyle(subHeaderStyle);
				count++;
			}
			Elements tds = row.select("td:not([rowspan])");
			count = 0;
			for (Element tdContent : tds) {
				// create cell for each
				cell = rowHeader.createCell(count);
				cell.setCellValue(tdContent.text());
				if(title=="Options (Packages)"){
					if(!pkgNames.contains(tdContent.firstElementSibling().text())){
						pkgNames.add(tdContent.firstElementSibling().text());
					}
				}
				count++;
			}
			rowCount++;
			//rowCount++; // get empty line after every content in excel
			// set auto size column for excel sheet
			sheet = wb.getSheetAt(0);
			for (int j = 0; j < row.select("th").size(); j++) {
				sheet.setColumnWidth(j, 8000);
				//sheet.autoSizeColumn(j);
			}
		}//rowCount++;   // one extra row after every table

	}

	@SuppressWarnings("deprecation")
	public void BYOExtExterior(XSSFWorkbook wb) {
		Connection.Response res;
		try {
			//Set Header Style
			XSSFCellStyle headerStyle = wb.createCellStyle();
			headerStyle.setFillForegroundColor(new XSSFColor(new java.awt.Color(22, 54, 92)));
			headerStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
			headerStyle.setAlignment(CellStyle.ALIGN_CENTER);

			//Set Sub-Header Style
			XSSFCellStyle subHeaderStyle = wb.createCellStyle();
			subHeaderStyle.setFillForegroundColor(new XSSFColor(new java.awt.Color(197, 217, 241)));
			subHeaderStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);

			//Set Header font color
			Font headerFont = wb.createFont();
			headerFont.setFontName("Calibri");
			headerFont.setBoldweight(Font.BOLDWEIGHT_NORMAL);
			headerFont.setColor(IndexedColors.WHITE.getIndex());
			headerStyle.setFont(headerFont);
			
			//Local Constants
			res = BYOExtractorConfig.confluenceConfig();
			Map<String, String> loginCookies  = res.cookies();
			XSSFSheet exteriorSheet = wb.createSheet("exterior");

			//Exterior section
			//System.out.println("Exterior section Started");
			Document exteriorPage = Jsoup.connect(BYOExtractorUtils.exteriorUrl).cookies(loginCookies).timeout(60000).get();

			/****************************Defaults********************************/

			defaults = exteriorPage.select("h1:contains(Defaults)").first();
			//System.out.println(defaults);
			defaultsTblContent = defaults.nextElementSibling();
			//System.out.println(defaultsTblContent);
			tblIterator(wb, exteriorSheet, "Defaults", defaultsTblContent);

			/****************************Paints********************************/
			paints = exteriorPage.select("h1:contains(Paints)").first();
			//System.out.println(paints);
			paintsTblcontent = paints.nextElementSibling();
			//System.out.println(paintsTblcontent);
			tblIterator(wb, exteriorSheet, "Paints", paintsTblcontent);

			/****************************Wheels********************************/
			wheels = exteriorPage.select("h1:contains(Wheels)").first();
			//System.out.println(wheels);
			wheelsTblContent = wheels.nextElementSibling();
			//System.out.println(wheelsTblContent);
			for(int k=0; k<=5; k++){
				//System.out.println(wheelsTblContent);
				if(wheelsTblContent.tagName().contains("div")){
					tblIterator(wb, exteriorSheet, "Wheels", wheelsTblContent);
					break;
				}else{
					wheelsTblContent = wheelsTblContent.nextElementSibling();
					//System.out.println(wheelsTblContent);
				}
			}

			/****************************(Options (Packages))********************************/
			optPkgs = exteriorPage.select("h1:contains(Options (Packages))").first();
			//System.out.println(optPkgs);
			optPkgsTblcontent = optPkgs.nextElementSibling();
			//System.out.println(optPkgsTblcontent);
			tblIterator(wb, exteriorSheet, "Options (Packages)", optPkgsTblcontent);

			for(String pkgNm : pkgNames){
				cell = exteriorSheet.createRow(rowCount).createCell(0);
				cell.setCellValue(pkgNm);
				Element nextsib = exteriorPage.select("h1:contains("+pkgNm+")").first();
				//System.out.println(nextsib.nextElementSibling());
				Element iteratePkgTblContent = nextsib.nextElementSibling();
				tblIterator(wb, exteriorSheet, pkgNm , iteratePkgTblContent);
			}

			/****************************Options (Factory Installed)********************************/
			facInstOpt = exteriorPage.select("h1:contains(Options (Factory Installed))").first();
			//System.out.println(rowCount);

			if(BYOExtractorUtils.verticalTable[0]){
				rowCount++;
				//exteriorSheet.addMergedRegion(new CellRangeAddress(rowCount-1, rowCount-1, 0, 4));
				cell = exteriorSheet.createRow(rowCount).createCell(0);
				exteriorSheet.addMergedRegion(new CellRangeAddress(rowCount, rowCount, 0, 4));
				cell.setCellValue("Options (Factory Installed)".toLowerCase());
				cell.setCellStyle(headerStyle);
				rowCount++;
				//System.out.println(rowCount);
				rowHeader = exteriorSheet.createRow(rowCount);
				cell = rowHeader.createCell(0);
				//cell.setCellValue("Options (Factory Installed)".toLowerCase());
				//cell.setCellStyle(headerStyle);
				//sheet.addMergedRegion(new CellRangeAddress(rowCount, rowCount, 0, 4));
				rowCount++;
				String[] headers = new String[] { "Name", "Image", "Image / Filename", "Copy", "Disclaimer", "Price", "Notes" };
				int count =0;
				//System.out.println(rowCount);
				//Row tempHeader = exteriorSheet.createRow(rowCount);
				for (String value : headers) {
					cell =  rowHeader.createCell(count);
					cell.setCellValue(value);
					cell.setCellStyle(subHeaderStyle);
					count++;
				}

				int l=0;
				optFactInstTblContent = facInstOpt.nextElementSibling();
				while(l<25){
					optFactInstTblContent = optFactInstTblContent.nextElementSibling();
					if(!optFactInstTblContent.tagName().contains("h1")){
						//System.out.println();
						if(!optFactInstTblContent.tagName().contains("p")){
							//	System.out.println(optFactInstTblContent);
							count =0;
							Row optInstHeader = exteriorSheet.createRow(rowCount);

							for (Element row : optFactInstTblContent.select("td")) {
								Elements tds = row.select("td:not([rowspan])");
								for (Element tdContent : tds) {
									cell = optInstHeader.createCell(count);
									cell.setCellValue(tdContent.text());
									count++;
								}
								exteriorSheet = wb.getSheetAt(0);
							}
							rowCount++;
						}else{
							continue;
						}
					}else{
						break;
					}
				}
			}else{
				tblIterator(wb, exteriorSheet, "Options (Factory Installed)", optPkgsTblcontent);
			}


			/****************************Accessories********************************/
			acc = exteriorPage.select("h1:contains(Accessories)").first();
			//System.out.println(acc);

			if(BYOExtractorUtils.verticalTable[0]){
				rowCount++;
				//exteriorSheet.addMergedRegion(new CellRangeAddress(rowCount, rowCount, 0, 4));
				cell = exteriorSheet.createRow(rowCount).createCell(0);
				exteriorSheet.addMergedRegion(new CellRangeAddress(rowCount, rowCount, 0, 4));
				cell.setCellValue("Accessories".toLowerCase());
				cell.setCellStyle(headerStyle);
				rowCount++;
				//System.out.println(rowCount);
				rowHeader = exteriorSheet.createRow(rowCount);
				cell = rowHeader.createCell(0);
				cell.setCellValue("Accessories".toLowerCase());
				cell.setCellStyle(headerStyle);
				//sheet.addMergedRegion(new CellRangeAddress(rowCount, rowCount, 0, 4));
				rowCount++;
				String[] headers = new String[] { "Name", "Image", "Image / Filename", "Copy", "Disclaimer", "Price", "Notes" };
				int count =0;
				for (String value : headers) {
					cell = rowHeader.createCell(count);
					cell.setCellValue(value);
					cell.setCellStyle(subHeaderStyle);
					count++;
				}
				int l=0;
				accTblContent = acc.nextElementSibling();
				while(l<25){
					accTblContent = accTblContent.nextElementSibling();
					if(accTblContent != null && accTblContent.id()!=null&&!accTblContent.id().contains("likes-and-labels-container") ){
						if(!accTblContent.tagName().contains("p")){
							//System.out.println(accTblContent);
							count = 0;
							Row accHeader = exteriorSheet.createRow(rowCount);
							for (Element row : accTblContent.select("td")) {
								Elements tds = row.select("td:not([rowspan])");
								for (Element tdContent : tds) {
									cell = accHeader.createCell(count);
									cell.setCellValue(tdContent.text());
									count++;
								}
								exteriorSheet = wb.getSheetAt(0);
							}
							rowCount++;
						}else{
							continue;
						}
					}else{
						break;
					}
				}
			}else{
				tblIterator(wb, exteriorSheet, "Accessories", accTblContent);
			}

		}catch(Exception e){
			e.printStackTrace();

		}
	}
}
