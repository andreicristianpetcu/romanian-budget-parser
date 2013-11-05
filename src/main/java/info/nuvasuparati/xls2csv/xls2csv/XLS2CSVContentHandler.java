package info.nuvasuparati.xls2csv.xls2csv;

import java.util.Arrays;
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
	
	@Override
	public void startElement(String uri, String localName, String name,
			Attributes atts) throws SAXException {
		tagStack.push(name);
		if (isInteresting(name)) {
			if(isSheet(name, atts)){
				
			}
			System.out.println(tagStack.toString() + ">>>" + name);
			if (atts instanceof AttributesImpl) {
				AttributesImpl attributesImpl = (AttributesImpl) atts;
				for (int i = 0; i < attributesImpl.getLength(); i++) {
					System.out.println("attributesValue="
							+ attributesImpl.getValue(i));
				}
			}
		}
		super.startElement(uri, localName, name, atts);
	}

	private boolean isSheet(String name, Attributes atts) {
		if(atts!=null && "div".equals(name) ){
			return atts.getValue("class").equals("page");
		}
		return false;
	}

	public void characters(char[] chars, int start, int length)
			throws SAXException {
		if (isInteresting(getLastTag())) {
			if(getLastTag().equals("h1")){
				System.out.println("------------------------------");
				System.out.println(new String(chars));
				System.out.println("------------------------------");
			} else {
				System.out.println(tagStack.toString() + "\t\t\t"
						+ new String(chars));
			}
		}
		super.characters(chars, start, length);
	}

	@Override
	public void endElement(String uri, String localName, String name)
			throws SAXException {
		tagStack.pop();
		if (isInteresting(name)) {
			System.out.println(tagStack.toString() + "<<<" + name);
		}
		super.endElement(uri, localName, name);
	}

	private boolean isInteresting(String name) {
		return interestingTags.contains(name);
	}

	private String getLastTag() {
		return tagStack.peek();
	}
}
