package com.marksheet.bps;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import com.itextpdf.kernel.colors.ColorConstants;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfPage;
import com.itextpdf.kernel.pdf.canvas.PdfCanvas;
import com.itextpdf.kernel.pdf.canvas.parser.listener.IPdfTextLocation;
import com.itextpdf.layout.Canvas;
import com.itextpdf.layout.element.Paragraph;
import com.itextpdf.layout.element.Text;
import com.itextpdf.pdfcleanup.PdfCleaner;
import com.itextpdf.pdfcleanup.autosweep.CompositeCleanupStrategy;
import com.itextpdf.pdfcleanup.autosweep.RegexBasedCleanupStrategy;
import com.marksheet.bps.constant.MarksheetConstant;

// TODO: Auto-generated Javadoc
/**
 * The Class PdfConverter.
 */
public class PdfConverter {

	/**
	 * The main method.
	 *
	 * @param args the arguments
	 * @throws IOException Signals that an I/O exception has occurred.
	 */
//	public static void main(String[] args) throws IOException {
//		PdfReader reader = new PdfReader("D:\\apps\\Sarvesh\\PetProject\\result-1st.pdf");
//		PdfWriter writer = new PdfWriter("D:\\apps\\Sarvesh\\PetProject\\result-2.pdf");
//		PdfDocument pdfDocument = new PdfDocument(reader, writer);
//		addContentToDocument(pdfDocument);
//		pdfDocument.close();
//
//	}
	public static void addDataToTheFile(String filePath, String targetFilePath, Map<Integer, String> data) {
		;
		List<String> paraList = new ArrayList<String>();
		try {
			XWPFDocument doc = new XWPFDocument(OPCPackage.open(new FileInputStream(filePath)));
			List<XWPFParagraph> paragraphList = doc.getParagraphs();
			for (XWPFParagraph para : paragraphList) {
				// if ((para.getStyle() != null) && (para.getNumFmt() != null)) {
				for (XWPFRun run : para.getRuns()) {
					String text = run.text();
					for (Entry<Integer, String> entry : data.entrySet()) {
						String key = "$" + entry.getKey();
						if (text.contains(key)) {
							text = text.replace(key, entry.getValue());
						}
					}
					run.setText(text, 0);
				}
				// }
				para.setAlignment(ParagraphAlignment.CENTER);
			}

			List<XWPFTable> tables = doc.getTables();
			for (XWPFTable tbl : doc.getTables()) {
				for (XWPFTableRow row : tbl.getRows()) {
					for (XWPFTableCell cell : row.getTableCells()) {

						String dataString = new String(cell.getText());
						dataString = dataString.strip();
						if (cell.getText() != null && cell.getText().contains("$")) {
							Integer keyData = null;
							String[] splitData = dataString.split("[$]");
							if (splitData.length == 2) {
								try {
									keyData = Integer.parseInt(dataString.replace(" ", "").replace("$", ""));
									dataString = dataString.replace(dataString, data.get(keyData)).replace(".0", "");
									cell.removeParagraph(0);
									cell.setText(dataString);

								} catch (Exception e) {
									System.err.println("Unable to parse : " + dataString);

								}
							}

//							else if (splitData.length > 2) {
//								String[] dataToProcess = dataString.split(" ");
//								for (int i = 0; i < cell.getParagraphs().size(); i++) {
//									cell.removeParagraph(i);
//								}
//								for (String stringdata : dataToProcess) {
//									stringdata = stringdata.strip();
//									if (stringdata.contains("$")) {
//										try {
//											keyData = Integer.parseInt(stringdata.replace(" ", "").replace("$", ""));
//											dataString = dataString.replace(stringdata, data.get(keyData)).replace(".0",
//													"");
//
//										} catch (Exception e) {
//											System.err.println("Unable to parse : " + stringdata);
//
//										}
//									}
//								}
//								for (int i = 0; i < cell.getParagraphs().size(); i++) {
//									cell.removeParagraph(i);
//								}
//
//								cell.setText(dataString);
//								System.out.println("sysData");
//							}

						}

						List<XWPFParagraph> paragraphList3 = cell.getParagraphs();
						for (XWPFParagraph para : cell.getParagraphs()) {
							para.setAlignment(ParagraphAlignment.CENTER);

						}
						System.out.println(cell.getText());

//						for (XWPFParagraph p : cell.getParagraphs()) {
//							for (XWPFRun r : p.getRuns()) {
//
//								String text = r.text();
//								if (text != null && text.contains("$")) {
//									Integer keyData = null;
//									try {
//										keyData = Integer.parseInt(text.replace(" ", "").replace("$", ""));
//										text = text.replace(text, data.get(keyData));
//										r.setText(text, 0);
//									} catch (Exception e) {
//										System.err.println("Unable to parse : " + text);
//
//									}
//
//								}
//							}
//						}
						for (int i = 0; i < cell.getParagraphs().size(); i++) {
							cell.getParagraphs().get(i).setAlignment(ParagraphAlignment.CENTER);
						}
					}
				}
			}

			doc.write(new FileOutputStream(targetFilePath));
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public static void main(String[] args) {

		// TODO Auto-generated method stub

		String filename = "D:\\apps\\Sarvesh\\PetProject\\result-1st_updated.docx";
		List<String> paraList = new ArrayList<String>();
		try {

			XWPFDocument doc = new XWPFDocument(OPCPackage.open(new FileInputStream(filename)));
			List<XWPFParagraph> paragraphList = doc.getParagraphs();
			for (XWPFParagraph para : paragraphList) {
				// if ((para.getStyle() != null) && (para.getNumFmt() != null)) {
				for (XWPFRun run : para.getRuns()) {
					String text = run.text();
					System.out.println(text);
					if (text.contains("Name of Student")) {
						System.out.println(text);
					}
					if (text.contains("Student___________________________________")) {
						text = text.replaceAll("Student___________________________________",
								"Student : ABHISHEK NAIDU      ");
					} else if (text.contains("Name______________________________")) {
						text = text.replaceAll("Name______________________________", "Name: BR Naidu      ");
						text = text.replaceAll("Class______", "Class : 1st      ");
					} else if (text.contains(" _____ ")) {
						text = text.replaceAll(" _____ ", " : A      ");
					} else if (text.contains("Year__________")) {
						text = text.replaceAll("Year__________", "Year : 2022-23      ");
					}

					run.setText(text, 0);
				}
				// }
			}
			doc.write(new FileOutputStream("D:\\apps\\Sarvesh\\PetProject\\result-1st_e.docx"));
		} catch (Exception e) {
			e.printStackTrace();
		}

//
//		Path fileName = Path.of("D:\\apps\\Sarvesh\\PetProject\\result-1st.doc");
//
//		BodyContentHandler handler = new BodyContentHandler();
//		Metadata metadata = new Metadata();
//		FileInputStream inputstream;
//		try {
//			ParseContext pcontext = new ParseContext();
//			inputstream = new FileInputStream(new File("D:\\apps\\Sarvesh\\PetProject\\result-1st.doc"));
//			AutoDetectParser parser = new AutoDetectParser();
//
//			parser.parse(inputstream, handler, metadata, pcontext);
//		} catch (FileNotFoundException e1) {
//			// TODO Auto-generated catch block
//			e1.printStackTrace();
//		} catch (IOException e) {
//			// TODO Auto-generated catch block
//			e.printStackTrace();
//		} catch (SAXException e) {
//			// TODO Auto-generated catch block
//			e.printStackTrace();
//		} catch (TikaException e) {
//			// TODO Auto-generated catch block
//			e.printStackTrace();
//		}
//
//		// OOXml parser
//
//		System.out.println("Contents of the document:" + handler.toString());
//		System.out.println("Metadata of the document:"+handler.);
//		String[] metadataNames = metadata.names();
//
//		for (String name : metadataNames) {
//			System.out.println(name + ": " + metadata.get(name));
//		}

		// Printing the string

//		OPCPackage pkg;
//		try {
//			pkg = OPCPackage.open(new File("D:\\\\apps\\\\Sarvesh\\\\PetProject\\\\result-1st.doc"));
//
//			POIXMLProperties props;
//
//			props = new POIXMLProperties(pkg);
//			props.getCoreProperties();
//			System.out.println("The title is " + props.getCoreProperties().getTitle());
//		} catch (IOException | OpenXML4JException | XmlException e) {
//			// TODO Auto-generated catch block
//			e.printStackTrace();
//		}

//		XWPFDocument doc = null;
////		XWPFDocument document = new XWPFDocument(new InputStream("D:\\apps\\Sarvesh\\PetProject\\result-1st.doc"));
//		try {
//			doc = new XWPFDocument(
//					OPCPackage.open(new FileInputStream("D:\\apps\\Sarvesh\\PetProject\\result-1st.doc")));
//		} catch (InvalidFormatException e) {
//			// TODO Auto-generated catch block
//			e.printStackTrace();
//		} catch (FileNotFoundException e) {
//			// TODO Auto-generated catch block
//			e.printStackTrace();
//		} catch (IOException e) {
//			// TODO Auto-generated catch block
//			e.printStackTrace();
//		}
//		TextReplacer textReplacer = new TextReplacer("Father's Name", "Sarvesh Choudhary");
//		textReplacer.replace(doc);
//		// TODO Auto-generated method stub

//		XWPFDocument doc = new XWPFDocument(OPCPackage.open("input.docx"));
//		for (XWPFParagraph p : doc.getParagraphs()) {
//			List<XWPFRun> runs = p.getRuns();
//			if (runs != null) {
//				for (XWPFRun r : runs) {
//					String text = r.getText(0);
//					if (text != null && text.contains("needle")) {
//						text = text.replace("needle", "haystack");
//						r.setText(text, 0);
//					}
//				}
//			}
//		}
//		for (XWPFTable tbl : doc.getTables()) {
//			for (XWPFTableRow row : tbl.getRows()) {
//				for (XWPFTableCell cell : row.getTableCells()) {
//					for (XWPFParagraph p : cell.getParagraphs()) {
//						for (XWPFRun r : p.getRuns()) {
//							String text = r.getText(0);
//							if (text != null && text.contains("needle")) {
//								text = text.replace("needle", "haystack");
//								r.setText(text, 0);
//							}
//						}
//					}
//				}
//			}
//		}
//		doc.write(new FileOutputStream("output.docx"));

	}

	/**
	 * Adds the content to document.
	 *
	 * @param pdfDocument the pdf document
	 * @throws IOException Signals that an I/O exception has occurred.
	 */
	private static void addContentToDocument(PdfDocument pdfDocument) throws IOException {

		// updateValuesInMarksheet(pdfDocument, MarksheetConstant.FIRST_STUDENT_NAME,
		// MarksheetConstant.FIRST_STUDENT_NAME_KEY, "Sarvesh Choudhary");
		updateValuesInMarksheet(pdfDocument, MarksheetConstant.FIRST_FATHER_NAME,
				MarksheetConstant.FIRST_FATHER_NAME_KEY, "Kavi Singh Choudhary");

		updateValuesInMarksheet(pdfDocument, MarksheetConstant.FIRST_CLASS_NAME, MarksheetConstant.FIRST_CLASS_NAME_KEY,
				"1st");
//		updateValuesInMarksheet(pdfDocument, MarksheetConstant.FIRST_SECTION_NAME,
//				MarksheetConstant.FIRST_SECTION_NAME_KEY, "A");
	}

	/**
	 * Update values in marksheet.
	 *
	 * @param pdfDocument the pdf document
	 * @param key         the key
	 * @param textKey     the text key
	 * @param value       the value
	 * @throws IOException Signals that an I/O exception has occurred.
	 */
	private static void updateValuesInMarksheet(PdfDocument pdfDocument, String key, String textKey, String value)
			throws IOException {
		CompositeCleanupStrategy strategy = new CompositeCleanupStrategy();
		strategy.add(new RegexBasedCleanupStrategy(key).setRedactionColor(ColorConstants.WHITE));
		PdfCleaner.autoSweepCleanUp(pdfDocument, strategy);

		for (IPdfTextLocation location : strategy.getResultantLocations()) {
			PdfPage page = pdfDocument.getPage(location.getPageNumber() + 1);
			PdfCanvas pdfCanvas = new PdfCanvas(page.newContentStreamAfter(), page.getResources(), page.getDocument());
			// PdfCanvas canvas = new PdfCanvas(page.newContentStreamAfter(),
			// page.getResources(), doc);
			Canvas canvas = new Canvas(pdfCanvas, location.getRectangle());
			String data = new Text(textKey).setBold().setFontSize(9).getText() + value;
			canvas.add(new Paragraph(data).setFontSize(9).setBold());
		}
	}

}