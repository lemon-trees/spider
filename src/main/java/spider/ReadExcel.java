package spider;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

public class ReadExcel {


	static String file_path = "D://D2018/邓田园简易版.xlsx";
	static String word_file_path = "D://D2018/图片.docx";
	static String file_path_new = "D://D2018/paixu/";
	static String default_file = "D://D2018/default.jpg";

	public static void main(String[] args) throws IOException, InvalidFormatException {

		InputStream inputStream = new FileInputStream(file_path);
		XSSFWorkbook wb = new XSSFWorkbook(inputStream);
		// 第二步，在workbook中添加一个sheet,对应Excel文件中的sheet
		XSSFSheet sheet = wb.getSheetAt(0);
		XSSFRow hssfRow;

		XWPFDocument doc = new XWPFDocument();
		XWPFTable table = doc.createTable((sheet.getLastRowNum() + 1), 2);
		List<String> pic = new ArrayList<>();
		for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
			hssfRow = sheet.getRow(rowIndex);
			if (hssfRow != null) {
				String key = null;
				XSSFCell cell = hssfRow.getCell(9);
				if (cell != null) {
					key = cell.getStringCellValue();
				}
				File oldfile = new File(key);
				if (oldfile.exists()) {
					pic.add(key);
				} else {
					pic.add(default_file);
				}
			}
		}
		int jishuIndex = 0;
		int oushuIndex = 0;
		for (int rowIndex = 0; rowIndex < pic.size(); rowIndex++) {
			System.out.println(pic.get(rowIndex));
			int picNo = rowIndex+1;
			File oldfile = new File(pic.get(rowIndex));
			FileInputStream input = new FileInputStream(oldfile);
			if (rowIndex % 2 == 0) {
				XWPFTableCell wordCell1 = table.getRow(jishuIndex).getCell(0);
				List<XWPFParagraph> paragraphs = wordCell1.getParagraphs();
				XWPFParagraph newPara = paragraphs.get(0);
				XWPFRun imageCellRunn = newPara.createRun();
				imageCellRunn.addPicture(input, XWPFDocument.PICTURE_TYPE_JPEG, oldfile.getName(), Units.toEMU(200),
						Units.toEMU(200)); // 200x200 pixels
				jishuIndex = jishuIndex+1;
				XWPFTableCell wordCelltu = table.getRow(jishuIndex).getCell(0);
				wordCelltu.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
				wordCelltu.setText("图片" + picNo);
				jishuIndex = jishuIndex+1;
			}else{
				XWPFTableCell wordCell2 = table.getRow(oushuIndex).getCell(1);
				List<XWPFParagraph> paragraphs = wordCell2.getParagraphs();
				XWPFParagraph newPara = paragraphs.get(0);
				XWPFRun imageCellRunn = newPara.createRun();
				imageCellRunn.addPicture(input, XWPFDocument.PICTURE_TYPE_JPEG, oldfile.getName(), Units.toEMU(200),
						Units.toEMU(200)); // 200x200 pixels
				oushuIndex = oushuIndex+1;
				XWPFTableCell wordCelltu = table.getRow(oushuIndex).getCell(1);
				wordCelltu.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
				wordCelltu.setText("图片" + picNo);
				oushuIndex = oushuIndex+1;
			}
			//OutputStream fileOutputStream = new FileOutputStream(newName);
			//byte[] b = new byte[input.available()];
			//input.read(b);
			//fileOutputStream.write(b);
			//input.close();
			//fileOutputStream.close();
			//wordCell.setText(rowIndex+"");





		}
		OutputStream wordOut = new FileOutputStream(word_file_path);
		doc.write(wordOut);
		wordOut.flush();
		wordOut.close();


	}

}
