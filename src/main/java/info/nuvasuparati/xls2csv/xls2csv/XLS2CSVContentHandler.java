package info.nuvasuparati.xls2csv.xls2csv;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Set;
import java.util.Stack;
import java.util.TreeSet;

import org.apache.tika.sax.BodyContentHandler;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.AttributesImpl;

public class XLS2CSVContentHandler extends BodyContentHandler {

	private Stack<String> tagStack = new Stack<String>();

	private Set<String> interestingTags = new TreeSet<String>(Arrays.asList(
			"tr", "td", "div", "h1"));

	private String currentSheetId;
	private String currentSheetName;
	private boolean insideTable = false;
	private List<String> currentRow = new ArrayList<String>();

	@Override
	public void startElement(String uri, String localName, String name,
			Attributes atts) throws SAXException {
		tagStack.push(name);
		super.startElement(uri, localName, name, atts);
	}

	public void characters(char[] chars, int start, int length)
			throws SAXException {
		String str = new String(chars).trim();
		if (isInteresting(getLastTag())) {
			computeCurrentSheetNameIfMissing(str);
			computeCurrentSheetIdIfMissing(str);
			computeInsideTableIfMissing(str);
			addToCurrentRowIfInsideTable(str);
		}
		super.characters(chars, start, length);
	}

	@Override
	public void endElement(String uri, String localName, String name)
			throws SAXException {
		tagStack.pop();
		if (isNonEmptyRowEnd(name)) {
			printEmptyRowEndedRowIsHeader();
			printEndedRowAndReset();
		}
		exitTableIfNecessary(name);
		super.endElement(uri, localName, name);
	}

	private void exitTableIfNecessary(String name) {
		if (name.equals("tbody")) {
			insideTable = false;
		}
	}

	private void printEndedRowAndReset() {
		System.out.println(currentRow.size() + " >> " + currentRow);
		currentRow.clear();
	}

	private void printEmptyRowEndedRowIsHeader() {
		if (isNewHeader()) {
			System.out.println();
		}
	}

	private boolean isNewHeader() {
		return currentRow.get(0).equals("Capitol");
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
			currentSheetName = null;
		}
	}

	private void computeCurrentSheetNameIfMissing(String str) {
		if (!str.trim().equals("")
				&& "[html, body, div, table, tbody, tr, td]".equals(tagStack
						.toString())) {
			if (currentSheetId != null && currentSheetName == null) {
				currentSheetName = str;
				System.out.println("------------------------------");
				System.out.println(currentSheetId + " - " + currentSheetName);
				System.out.println("------------------------------");
			}
		}
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
