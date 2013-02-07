/*
 * Copyright 2013 Future Systems
 * 
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * 
 * http://www.apache.org/licenses/LICENSE-2.0
 * 
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package org.krakenapps.docxcod;

import static org.krakenapps.docxcod.util.XMLDocHelper.evaluateXPath;
import static org.krakenapps.docxcod.util.XMLDocHelper.newXPath;

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
