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
import static org.krakenapps.docxcod.util.XMLDocHelper.evaluateXPathExpr;
import static org.krakenapps.docxcod.util.XMLDocHelper.newDocumentBuilder;
import static org.krakenapps.docxcod.util.XMLDocHelper.newXPath;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.xml.transform.TransformerFactoryConfigurationError;
import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathExpression;
import javax.xml.xpath.XPathExpressionException;

import org.krakenapps.docxcod.util.CloseableHelper;
import org.krakenapps.docxcod.util.XMLDocHelper;
import org.krakenapps.docxcod.util.XMLDocHelper.NodeListIterAdapter;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

public class MergeFieldParser implements OOXMLProcessor {

	private Logger logger = LoggerFactory.getLogger(getClass().getName());

	public void process(OOXMLPackage pkg, Map<String, Object> rootMap) {
		/*
		 * extract contents from merge fields and make magic node containing
		 * them in proper position.
		 */
		extractMergeField(pkg);
	}

	private void extractMergeField(OOXMLPackage pkg) throws TransformerFactoryConfigurationError {
		InputStream f = null;
		try {
			f = new FileInputStream(new File(pkg.getDataDir(), "word/document.xml"));
			Document doc = newDocumentBuilder().parse(f);

			XPath xpath = newXPath(doc);
			NodeList nodeList = evaluateXPath(xpath,
					"//*[name()='w:fldChar' or name()='w:instrText' or name()='w:fldSimple']", doc);

			List<Directive> directives = DirectiveExtractor.parseNodeList(nodeList);
			for (Directive d : directives) {

				Node n = d.getPosition();

				String directive = d.getDirectiveString();
				logger.debug("{} {}", new Object[] { n.getNodeName(), directive });

				MakeMagicNode(doc, n, directive);
			}

			XMLDocHelper.save(doc, new File(pkg.getDataDir(), "word/document.xml"), true);

		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			CloseableHelper.safeClose(f);
		}
	}

	private void MakeMagicNode(Document doc, Node n, String directive) {
		/*
		 * move all nodes in fldSimple to out of it. and replace text contents
		 * of <w:t> with KMagicNode
		 */
		if (n.getNodeName().equals("w:fldSimple")) {
			/* // @formatter:off
            <w:fldSimple w:instr=" MERGEFIELD  @after-row/#list  \* MERGEFORMAT ">
              <w:r w:rsidR="00771959">
                <w:rPr>
                  <w:noProof/>
                </w:rPr>
                <w:t>«@after-row/</w:t>
              </w:r>
              <w:r w:rsidR="00771959">
                <w:rPr>
                  <w:rFonts w:hint="eastAsia"/>
                  <w:noProof/>
                </w:rPr>
                <w:t>#</w:t>
              </w:r>
              <w:r w:rsidR="00771959">
                <w:rPr>
                  <w:noProof/>
                </w:rPr>
                <w:t>list»</w:t>
              </w:r>
            </w:fldSimple>


            */ // @formatter:on
			logger.debug("fldSimple found");
			XPath xpath = newXPath(doc);
			XPathExpression xpFldSimpleText;
			Node parent = n.getParentNode();
			try {
				xpFldSimpleText = xpath.compile("w:r/w:t");
				NodeList t = evaluateXPathExpr(xpFldSimpleText, n);

				t.item(0).setTextContent("");
				t.item(0).appendChild(getMagicNode(doc, directive));
				parent.insertBefore(t.item(0).getParentNode().cloneNode(true), n);
			} catch (XPathExpressionException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}

			// n : w:fldSimple can contain many w:r in its children
			// for (Node c : new NodeListWrapper(n.getChildNodes())) {
			// if (c.getNodeName() != null)
			// parent.insertBefore(c.cloneNode(true), n);
			// }
			parent.removeChild(n);
		} else if (n.getNodeName().equals("w:fldChar")) {
			logger.debug("fldChar found");
			// @formatter:off
			/*
            <w:r>
              <w:fldChar w:fldCharType="begin" />
            </w:r>
            <w:r>
              <w:instrText xml:space="preserve">MERGEFIELD @after-row#/list \* MERGEFORMAT</w:instrText>
            </w:r>
            <w:r>
              <w:fldChar w:fldCharType="separate" />
            </w:r>
            <w:r> <!-- style of this run will be used -->
              <w:rPr>
                <w:noProof />
              </w:rPr>
              <w:t>«@after-row#/list»</w:t>
            </w:r>
            <w:r>
              <w:rPr>
                <w:noProof />
              </w:rPr>
              <w:fldChar w:fldCharType="end" />
            </w:r>
			 */
            // @formatter:on
			Node firstRun = n.getParentNode();
			Node beginRun = null;
			Node sibling = firstRun;
			Node lastRun = null;
			Node newRun = null;
			ArrayList<Node> willBeRemoved = new ArrayList<Node>();

			while (sibling != null)
			{
				if (sibling.getNodeName().equals("w:r"))
				{
					// all nodes in w:fldChar except 'newRun' will be removed
					// finally.
					willBeRemoved.add(sibling);

					Node fldCharNode = findFldCharNode(sibling);
					if (fldCharNode == null) {
						sibling = sibling.getNextSibling();
						continue;
					}
					NamedNodeMap attributes = fldCharNode.getAttributes();
					Node namedItem = attributes.getNamedItem("w:fldCharType");
					if (namedItem == null) {
						sibling = sibling.getNextSibling();
						continue;
					}
					if (namedItem.getNodeValue().equals("begin")) {
						// beginRun will not be removed
						beginRun = sibling;

						sibling = sibling.getNextSibling();

						continue;// live
					}
					if (namedItem.getNodeValue().equals("separate")) {
						// beginRun will not be removed
						beginRun = sibling;

						sibling = sibling.getNextSibling();

						// skip whitespace elements and find first w:r.
						while (beginRun != null && !beginRun.getNodeName().equals("w:r")) {
							beginRun = beginRun.getNextSibling();
						}

						continue;// live
					}
					if (namedItem.getNodeValue().equals("end")) {
						lastRun = sibling;

						break;
					}
				}
				sibling = sibling.getNextSibling();
			}
			willBeRemoved.remove(beginRun);

			if (beginRun != null && lastRun != null) { // found matching "end" fldChar
				do {
					Node parentNode = firstRun.getParentNode();
					for (Node node : willBeRemoved) {
						parentNode.removeChild(node);
					}

					// replace contents of first w:r. so formating style of
					// newRun will preserved.
					Node textNode = findTextNode(beginRun);
					if (textNode != null) {
						textNode.setTextContent("");
						textNode.appendChild(getMagicNode(doc, directive));
						break;
					}
					Node fldCharNode = findNode(beginRun, "w:fldChar"); 
					if (fldCharNode != null) {
						textNode = doc.createElement("w:t");
						textNode.appendChild(getMagicNode(doc, directive));
						beginRun.appendChild(textNode);
						beginRun.removeChild(fldCharNode);
						break;
					}
					
					logger.warn("not-supported fldChar type node");

				} while (false);
			} else {
				logger.warn("no coressponding begin & end fldChar found");
			}
		}

	}

	private static Pattern MAGICNODE_PATTERN = Pattern.compile("<KMagicNode><![CDATA[+(.*)+]]></KMagicNode>");

	private static String parseMagicNode(String in) {
		in = replaceUnicodeQuote(in.trim());
		Matcher matcher = MAGICNODE_PATTERN.matcher(in);

		if (matcher.find() && matcher.groupCount() > 0) {
			String f = matcher.group(1);
			if (f == null)
				f = matcher.group(2);
			f = f.replaceAll("\\\\(.)", "$1");
			return f;
		} else
			return null;
	}

	private static String replaceUnicodeQuote(String in) {
		StringBuilder builder = new StringBuilder();
		for (int i = 0; i < in.length(); ++i) {
			int type = Character.getType(in.codePointAt(i));
			switch (type) {
			case Character.FINAL_QUOTE_PUNCTUATION:
			case Character.INITIAL_QUOTE_PUNCTUATION:
				builder.append('"');
				break;
			default:
				builder.append(in.charAt(i));
				break;
			}
		}
		return builder.toString();
	}

	// unused but leave here for reference
	@SuppressWarnings("unused")
	private static String transformMagicNode(String nodeValue) {
		String result = nodeValue.trim();
		if (result.startsWith("<KMagicNode>")) {
			result = parseMagicNode(result);
		}
		return result;
	}

	public static final String UTF8_BOM = "\uFEFF";

	private Node findFldCharNode(Node sibling) {
		for (Node n : new NodeListIterAdapter(sibling.getChildNodes())) {
			if (n.getNodeName().equals("w:fldChar")) {
				return n;
			}
		}
		return null;
	}

	private Node findNode(Node sibling, String elemName) {
		for (Node n : new NodeListIterAdapter(sibling.getChildNodes())) {
			if (n.getNodeName().equals(elemName)) {
				return n;
			}
		}
		return null;
	}

	private Node findTextNode(Node sibling) {
		if (sibling.getFirstChild() == null)
			return null;
		for (Node n : new NodeListIterAdapter(sibling.getChildNodes())) {
			if (n.getNodeName().equals("w:t")) {
				return n;
			}
		}
		return null;
	}

	private Node getMagicNode(Document doc, String content) {
		Element magicNode = doc.createElement("KMagicNode");
		magicNode.appendChild(doc.createCDATASection(content));
		return magicNode;
	}
}
