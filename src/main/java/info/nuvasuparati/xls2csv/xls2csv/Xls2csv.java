package info.nuvasuparati.xls2csv.xls2csv;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

import org.apache.tika.Tika;
import org.apache.tika.exception.TikaException;
import org.apache.tika.metadata.Metadata;
import org.apache.tika.parser.AutoDetectParser;
import org.apache.tika.parser.ParseContext;
import org.apache.tika.sax.BodyContentHandler;

public class Xls2csv {

	public static void main(String[] args) throws FileNotFoundException,
			IOException, org.xml.sax.SAXException, TikaException {

		InputStream is = Xls2csv.class
				.getResourceAsStream("/ANEXAOPC_SINTEZA_Buget2013.xls");
		
		System.out.println("is=" + is);
		Metadata metadata = new Metadata();
		BodyContentHandler ch = new BodyContentHandler();
		AutoDetectParser parser = new AutoDetectParser();

		String mimeType = new Tika().detect(is);
		metadata.set(Metadata.CONTENT_TYPE, mimeType);

		parser.parse(is, ch, metadata, new ParseContext());
		is.close();

		for (int i = 0; i < metadata.names().length; i++) {
			String item = metadata.names()[i];
			System.out.println(item + " -- " + metadata.get(item));
		}

		System.out.println(ch.toString());
	}
}