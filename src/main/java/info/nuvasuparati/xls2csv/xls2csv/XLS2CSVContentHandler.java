package info.nuvasuparati.xls2csv.xls2csv;

import java.io.File;
import java.io.FileWriter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Set;
import java.util.Stack;
import java.util.TreeSet;

import org.apache.tika.sax.BodyContentHandler;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;

import au.com.bytecode.opencsv.CSVWriter;

public class XLS2CSVContentHandler extends BodyContentHandler {

	private Stack<String> tagStack = new Stack<String>();

	private Set<String> interestingTags = new TreeSet<String>(Arrays.asList(
			"tr", "td", "div", "h1"));

	private String currentSheetId;
	private String currentSheetName;
	private boolean insideTable = false;
	private List<String> currentRow = new ArrayList<String>();
	private List<String> rowHeader = new ArrayList<String>();

	private boolean tdWithNoText;

	private File currentCSVFile;

	private CSVWriter csvWriter;

	@Override
	public void startElement(String uri, String localName, String tagName,
			Attributes atts) throws SAXException {
		tagStack.push(tagName);
		resetTdWithNoTextFlag(tagName);
		super.startElement(uri, localName, tagName, atts);
	}

	public void characters(char[] chars, int start, int length)
			throws SAXException {
		String str = new String(chars).trim();
		String lastTag = getLastTag();
		if (isInteresting(lastTag)) {
			computeCurrentSheetNameIfMissing(str);
			computeCurrentSheetIdIfMissing(str);

			computeInsideTableIfMissing(str);

			if (isStartOfTableHeader(str)) {
				currentRow.clear();
			}

			addToCurrentRowIfInsideTable(str);
			markIfTdHasText(lastTag);
		}
		super.characters(chars, start, length);
	}

	@Override
	public void endElement(String uri, String localName, String name)
			throws SAXException {
		String currentTag = tagStack.pop();
		addEmptyColumnIfNeeded(currentTag);
		if (isNonEmptyRowEnd(name)) {
			printEndedRowAndReset();
		}
		exitTableIfNecessary(name);

		super.endElement(uri, localName, name);
	}

	private void addEmptyColumnIfNeeded(String currentTag) {
		if ("td".equals(currentTag) && tdWithNoText) {
			currentRow.add("");
		}
	}

	private void resetTdWithNoTextFlag(String tagName) {
		if ("td".equals(tagName)) {
			tdWithNoText = true;
		}
	}

	private void markIfTdHasText(String lastTag) {
		if ("td".equals(lastTag)) {
			tdWithNoText = false;
		}
	}

	private void exitTableIfNecessary(String name) {
		if (name.equals("tbody")) {
			insideTable = false;
			finishWriteCsv();
		}
	}

	private void startWritingCsvFile() {
		currentCSVFile = new File(getFullSheetName() + ".csv");
		System.out.println(">>" + currentCSVFile.getAbsolutePath());
		try {
			FileWriter fileWriter = new FileWriter(currentCSVFile);
			csvWriter = new CSVWriter(fileWriter, '\t');
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	private void finishWriteCsv() {
		try {
			csvWriter.close();
		} catch (Exception e) {
			throw new RuntimeException();
		}
	}

	private void printEndedRowAndReset() {
		boolean currentRowHeader = isCurrentRowHeader();
		if (currentRowHeader) {
			rowHeader = new ArrayList<String>(currentRow);
		}
		ArrayList<String> rowToPrint = new ArrayList<String>();
		for (int i = 0; i < rowHeader.size(); i++) {
			if (!rowHeader.get(i).trim().equals("")) {
				rowToPrint.add(currentRow.get(i));
			}
		}

		System.out.println(rowToPrint);
		csvWriter.writeNext(rowToPrint.toArray(new String[0]));
		currentRow.clear();
	}

	private boolean isCurrentRowHeader() {
		boolean currentRowHeader = currentRow.size() > 0
				&& currentRow.get(0).trim().equals("Capitol");
		return currentRowHeader;
	}

	private boolean isNonEmptyRowEnd(String name) {
		return insideTable && name.equals("tr") && currentRow.size() > 0;
	}

	private void addToCurrentRowIfInsideTable(String str) {
		if (insideTable) {
			currentRow.add(str);
		}
	}

	private void computeInsideTableIfMissing(String str) {
		if (isStartOfTableHeader(str)) {
			insideTable = true;
		}
	}

	private void computeCurrentSheetIdIfMissing(String str) {
		if (getLastTag().equals("h1") && str != null && !str.trim().equals("")) {
			currentSheetId = str;
			currentSheetName = "";
		}
	}

	private void computeCurrentSheetNameIfMissing(String str) {
		if ("Bugetul pe anul 2013".equals(str.trim())) {
			System.out.println();
			System.out.println("------------------------------");
			System.out.println(getFullSheetName());
			System.out.println("------------------------------");
			startWritingCsvFile();
			currentSheetId = null;
			currentSheetName = null;
		}
		if (!str.trim().equals("")
				&& "[html, body, div, table, tbody, tr, td]".equals(tagStack
						.toString())) {
			if (currentSheetId != null && currentSheetName != null) {
				currentSheetName += (!currentSheetName.equals("") ? " " : "")
						+ str;
			}
		}
	}

	private String getFullSheetName() {
		return (currentSheetId + "_" + currentSheetName).replace("-", "_")
				.replace(" ", "_").replace("___", "_").replace("__", "_");
	}

	private boolean isSheet(String name, Attributes atts) {
		if (atts != null && "div".equals(name)) {
			return atts.getValue("class").equals("page");
		}
		return false;
	}

	private boolean isStartOfTableHeader(String str) {
		boolean isHeaderFirstColumn = str != null
				&& str.trim().equals("Capitol");
		boolean isTd = "td".equals(getLastTag());
		return isHeaderFirstColumn && isTd;
	}

	private boolean isInteresting(String name) {
		return interestingTags.contains(name);
	}

	private String getLastTag() {
		return tagStack.peek();
	}
}
