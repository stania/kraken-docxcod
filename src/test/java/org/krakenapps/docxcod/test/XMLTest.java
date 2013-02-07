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
package org.krakenapps.docxcod.test;

import static org.junit.Assert.assertTrue;
import static org.krakenapps.docxcod.util.XMLDocHelper.evaluateXPath;
import static org.krakenapps.docxcod.util.XMLDocHelper.evaluateXPathExpr;
import static org.krakenapps.docxcod.util.XMLDocHelper.newDocumentBuilder;
import static org.krakenapps.docxcod.util.XMLDocHelper.newXPath;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;

import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathExpression;
import javax.xml.xpath.XPathExpressionException;

import org.custommonkey.xmlunit.Diff;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import org.krakenapps.docxcod.util.XMLDocHelper.NodeListIterAdapter;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.InputSource;

public class XMLTest {
	private TearDownHelper tearDownHelper = new TearDownHelper();

	@Before
	public void setUp() throws Exception {
	}

	@After
	public void tearDown() throws Exception {
		tearDownHelper.tearDown();
	}

	@Test
	public void xmlTest2() throws Exception {
		InputStream f = getClass().getResourceAsStream("/sampledoc/word/document.xml");
		assertTrue(f != null);

		{
			Document doc = newDocumentBuilder().parse(f);
			assertTrue(doc != null);

			XPath xpath = newXPath(doc);

			NodeList nodeList = evaluateXPath(xpath, "/w:document[1]/w:body[1]/w:p[2]/w:r[1]/w:t[1]", doc);
			for (Node node : new NodeListIterAdapter(nodeList)) {
				Node targetPara = node.getParentNode().getParentNode();
				Node parentOfPara = targetPara.getParentNode();
				parentOfPara.insertBefore(getMagicNode(doc), targetPara);
				parentOfPara.insertBefore(getMagicNode(doc), targetPara.getNextSibling());
			}

			Transformer transformer = TransformerFactory.newInstance().newTransformer();
			transformer.setOutputProperty(OutputKeys.INDENT, "yes");
			transformer.setOutputProperty(OutputKeys.CDATA_SECTION_ELEMENTS, "yes");
			transformer.setOutputProperty("{http://xml.apache.org/xslt}indent-amount", "2");
			transformer.transform(new DOMSource(doc), new StreamResult("test2.xml"));
			tearDownHelper.add(new File("test2.xml"));
		}
		{
			Document genDoc = newDocumentBuilder().parse(new File("test2.xml"));
			XPath xpath = newXPath(genDoc);
			checkNodeName(genDoc, xpath, "/w:document[1]/w:body[1]/*[2]", "KMagicNode");
			checkNodeName(genDoc, xpath, "/w:document[1]/w:body[1]/*[3]", "w:p");
			checkNodeName(genDoc, xpath, "/w:document[1]/w:body[1]/*[4]", "KMagicNode");
		}
	}

	private void checkNodeName(Document genDoc, XPath xpath, String expr, String nodeName)
			throws XPathExpressionException {
		NodeList nodeList = evaluateXPath(xpath, expr, genDoc);
		assertTrue(nodeList.getLength() > 0);
		assertTrue(nodeList.item(0).getNodeName().equals(nodeName));
	}

	private Node getMagicNode(Document doc) {
		Element magicNode = doc.createElement("KMagicNode");
		magicNode.appendChild(doc.createCDATASection("some template engine syntax"));
		return magicNode;
	}

}
