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
		// if (isInteresting(name)) {
		// // System.out.println(tagStack.toString() + ">>>" + name);
		// if (atts instanceof AttributesImpl) {
		// AttributesImpl attributesImpl = (AttributesImpl) atts;
		// for (int i = 0; i < attributesImpl.getLength(); i++) {
		// System.out.println("attributesValue="
		// + attributesImpl.getValue(i));
		// }
		// }
		// }
		super.startElement(uri, localName, name, atts);
	}

	public void characters(char[] chars, int start, int length)
			throws SAXException {
		String str = new String(chars).trim();
		if (isInteresting(getLastTag())) {
			if (!str.trim().equals("")
					&& "[html, body, div, table, tbody, tr, td]"
							.equals(tagStack.toString())) {
				if (currentSheetId != null && currentSheetName == null) {
					currentSheetName = str;
					System.out.println("------------------------------");
					System.out.println(currentSheetId + " - "
							+ currentSheetName);
					System.out.println("------------------------------");
				}
			}
			if (getLastTag().equals("h1") && str!=null && !str.trim().equals("")) {
				currentSheetId = str;
				currentSheetName = null;
			} else if (isStartOfTableHeader(str)) {
				insideTable = true;
			} else {
				// System.out.println(tagStack.toString() + "\t\t\t"
				// + currentSheetId);
			}
			if (insideTable) {
				currentRow.add(str);
			}
		}
		super.characters(chars, start, length);
	}

	@Override
	public void endElement(String uri, String localName, String name)
			throws SAXException {
		tagStack.pop();
		if (isInteresting(name)) {
			// System.out.println(tagStack.toString() + "<<<" + name);
		}
		if (name.equals("tr") && insideTable && currentRow.size() > 0) {
			if (currentRow.get(0).equals("Capitol")) {
				System.out.println();
			}
			System.out.println(currentRow);
			currentRow.clear();
		}
		if (name.equals("tbody")) {
			insideTable = false;
		}
		super.endElement(uri, localName, name);
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
