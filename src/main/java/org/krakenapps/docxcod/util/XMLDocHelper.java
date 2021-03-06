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
package org.krakenapps.docxcod.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Iterator;

import javax.xml.namespace.NamespaceContext;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.TransformerFactoryConfigurationError;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathConstants;
import javax.xml.xpath.XPathExpression;
import javax.xml.xpath.XPathExpressionException;
import javax.xml.xpath.XPathFactory;

import org.krakenapps.docxcod.OOXMLPackage;
import org.w3c.dom.DOMException;
import org.w3c.dom.Document;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import static org.krakenapps.docxcod.util.CloseableHelper.safeClose;

public class XMLDocHelper {
	static private DocumentBuilderFactory dbFactory;
	static private XPathFactory xpathFactory;

	public static DocumentBuilder newDocumentBuilder() throws ParserConfigurationException {
		if (dbFactory == null) {
			dbFactory = DocumentBuilderFactory.newInstance();
			dbFactory.setNamespaceAware(true);
		}
		return dbFactory.newDocumentBuilder();
	}

	public static XPath newXPath(Document doc) {
		if (xpathFactory == null) {
			xpathFactory = XPathFactory.newInstance();
		}

		XPath xpath = xpathFactory.newXPath();
		xpath.setNamespaceContext(new DocumentNamespaceContext(doc));
		return xpath;
	}

	private static class DocumentNamespaceContext implements NamespaceContext {
		Document doc;
		String unnamedPrefix = null;
		String unnamedURI = null;

		HashMap<String, String> uriMap = new HashMap<String, String>();

		public DocumentNamespaceContext(Document doc) {
			this.doc = doc;
			if (doc.lookupNamespaceURI(null) != null) {
				unnamedPrefix = "DEF";
				unnamedURI = doc.lookupNamespaceURI(null);
			}
			try {
				XPath xpath = xpathFactory.newXPath();
				NodeList nodeList_ = evaluateXPath(xpath, "//*/namespace::*", doc);

				for (Node n : new NodeListIterAdapter(nodeList_)) {
					String nsName = n.getNodeName();
					String nsURI = n.getTextContent();
					if (nsName.length() > 6) {
						nsName = nsName.substring(6);
					} else {
						nsName = ""; // empty(default) namespace
					}
					uriMap.put(nsName, nsURI);
				}
				
			} catch (XPathExpressionException e) {
			}
		}

		@Override
		public String getNamespaceURI(String prefix) {
			if (prefix.equals(unnamedPrefix))
				return unnamedURI;
			if (uriMap.containsKey(prefix))
				return uriMap.get(prefix);
			else
				return doc.lookupNamespaceURI(prefix);
		}

		@Override
		public String getPrefix(String namespaceURI) {
			// no need to implement for using xpath
			return null;
		}

		@Override
		public Iterator<?> getPrefixes(String namespaceURI) {
			// no need to implement for using xpath
			return null;
		}
	}

	public static NodeList evaluateXPath(Document doc, String expression) throws XPathExpressionException {
		return (NodeList) newXPath(doc).evaluate(expression, doc, XPathConstants.NODESET);
	}

	public static NodeList evaluateXPath(XPath xpath, String expression, Node node) throws XPathExpressionException {
		return (NodeList) xpath.evaluate(expression, node, XPathConstants.NODESET);
	}

	public static NodeList evaluateXPathExpr(XPathExpression xpath, Node node) throws XPathExpressionException {
		return (NodeList) xpath.evaluate(node, XPathConstants.NODESET);
	}

	public static class NodeListIterAdapter implements Iterable<Node> {
		private final NodeList nl;

		public NodeListIterAdapter(NodeList nl) {
			this.nl = nl;
		}

		public class IteratorI implements Iterator<Node> {
			private int current;

			public IteratorI() {
				this.current = 0;
			}

			@Override
			public boolean hasNext() {
				if (nl == null)
					return false;
				else
					return current < nl.getLength();
			}

			@Override
			public Node next() {
				if (nl == null)
					return null;
				else
					return nl.item(current++);
			}

			@Override
			public void remove() {
			}
		}

		@Override
		public Iterator<Node> iterator() {
			return new IteratorI();
		}
	}

	public static void save(Document doc, File file, boolean indent) throws XPathExpressionException, DOMException,
			TransformerFactoryConfigurationError, TransformerException {
		doc.normalize();

		XPath xpath = newXPath(doc);

		for (Node n : new NodeListIterAdapter(evaluateXPath(xpath, "//text()[normalize-space(.)='']", doc))) {
			n.getParentNode().removeChild(n);
		}

		Transformer transformer = TransformerFactory.newInstance().newTransformer();
		transformer.setOutputProperty(OutputKeys.INDENT, indent ? "yes" : "no");
		transformer.setOutputProperty(OutputKeys.CDATA_SECTION_ELEMENTS, "yes");
		transformer.setOutputProperty("{http://xml.apache.org/xslt}indent-amount", "2");
		transformer.transform(new DOMSource(doc), new StreamResult(file));
	}
	
	public static Document parseXml(OOXMLPackage pkg, String path) {
		InputStream is = null;
		try {
			is = new FileInputStream(new File(pkg.getDataDir(), path));
			Document doc = newDocumentBuilder().parse(is);

			return doc;
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			safeClose(is);
		}

		return null;
	}

	public static String getNodeAttribute(Node n, String attrName) {
		Node namedAttr = n.getAttributes().getNamedItem(attrName);
		if (namedAttr != null)
			return namedAttr.getNodeValue();
		else
			return null;
	}

	public static void setNodeAttribute(Document doc, Node n, String attrName, String value) {
		Node namedAttr = n.getAttributes().getNamedItem(attrName);
		if (namedAttr == null) {
			namedAttr = doc.createAttribute(attrName);
			n.getAttributes().setNamedItem(namedAttr);
		}
		namedAttr.setNodeValue(value);
	}


}
