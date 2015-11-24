package com.java;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

public class ReplaceTextPOI {

	public static void readAndReplaceWord(String inputFile, String outputFile) {
		try {

			File file = new File(inputFile);

			if (!file.exists()) {
				throw new IOException("Word file does not exist!");
			}

			FileInputStream is = new FileInputStream(file);
			XWPFDocument doc = new XWPFDocument(is);
			List<XWPFParagraph> paragraphs = doc.getParagraphs();

			for (XWPFParagraph p : doc.getParagraphs()) {
				List<XWPFRun> runs = p.getRuns();
				if (runs != null) {
					for (XWPFRun r : runs) {
						String text = r.getText(0);
						if (text != null && !"".equals(text.trim())) {
							if (text.contains("##MESUT")) {
								text = text.replaceAll("##MESUT", "mesut");
								System.out.println(text);
								r.setText(text, 0);
							}
						}
					}
				}
			}

			for (XWPFTable tbl : doc.getTables()) {
				for (XWPFTableRow row : tbl.getRows()) {
					for (XWPFTableCell cell : row.getTableCells()) {
						for (XWPFParagraph p : cell.getParagraphs()) {
							for (XWPFRun r : p.getRuns()) {
								String text = r.getText(0);
								if (text != null && !"".equals(text.trim())) {
									if (text.contains("##MESUT")) {
										text = text.replaceAll("##MESUT", "mesut");
										System.out.println(text);
										r.setText(text, 0);
									}
								}

							}

						}
					}
				}
			}
			doc.write(new FileOutputStream(outputFile));
		} catch (Exception e) {
			e.printStackTrace();
		}

	}
}
