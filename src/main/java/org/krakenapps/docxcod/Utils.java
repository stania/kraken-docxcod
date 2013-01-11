package org.krakenapps.docxcod;

import static org.krakenapps.docxcod.util.XMLDocHelper.evaluateXPath;
import static org.krakenapps.docxcod.util.XMLDocHelper.newXPath;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathExpressionException;

import org.krakenapps.docxcod.util.XMLDocHelper.NodeListIterAdapter;
import org.w3c.dom.Document;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

public class Utils {

	public static List<String> readSharedStrings(Document sharedStringsDoc) throws XPathExpressionException {
		ArrayList<String> list = new ArrayList<String>();

		XPath xpath = newXPath(sharedStringsDoc);

		NodeList nodes = evaluateXPath(xpath, "//DEF:si/DEF:t", sharedStringsDoc);
		for (Node n : new NodeListIterAdapter(nodes)) {
			list.add(n.getTextContent());
		}

		return list;
	}

}
