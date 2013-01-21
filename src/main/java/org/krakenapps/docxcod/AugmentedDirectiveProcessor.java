package org.krakenapps.docxcod;

import static org.krakenapps.docxcod.util.CloseableHelper.safeClose;
import static org.krakenapps.docxcod.util.XMLDocHelper.evaluateXPath;
import static org.krakenapps.docxcod.util.XMLDocHelper.newDocumentBuilder;
import static org.krakenapps.docxcod.util.XMLDocHelper.newXPath;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.Map;

import javax.xml.xpath.XPath;

import org.krakenapps.docxcod.util.XMLDocHelper;
import org.krakenapps.docxcod.util.XMLDocHelper.NodeListIterAdapter;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

public class AugmentedDirectiveProcessor implements OOXMLProcessor {
	private Logger logger = LoggerFactory.getLogger(getClass().getName());

	static String[] augmentedDirectives = {
			"@before-row",
			"@after-row",
			"@row",
			"@/row",
			"@before-para",
			"@after-para",
			"@para",
			"@/para",
	};

	@SuppressWarnings("serial")
	private static class CannotParseAugmentedDirectiveException extends RuntimeException {
	}

	private static class AugmentedDirective {
		private static String[][] augmentedDirectives = {
				{ "@before-row", "before", "w:tr" },
				{ "@row", "before", "w:tr" },
				{ "@after-row", "after", "w:tr" },
				{ "@/row", "after", "w:tr" },
				{ "@before-para", "before", "w:p" },
				{ "@para", "before", "w:p" },
				{ "@after-para", "after", "w:p" },
				{ "@/para", "after", "w:p" },
		};
		private String prefix;
		private String remaining;
		private boolean isBefore;
		private String expectedParent;

		private AugmentedDirective(String augDirective, String rawDirective, boolean isBefore,
				String expectedParentElementName) {
			this.prefix = augDirective;
			this.remaining = rawDirective.substring(augDirective.length());
			this.isBefore = isBefore;
			this.expectedParent = expectedParentElementName;
		}

		public static AugmentedDirective parseDirective(String directive) {

			if (!directive.startsWith("@"))
				throw new CannotParseAugmentedDirectiveException();

			for (String[] entry : augmentedDirectives) {
				String prefix = entry[0];

				if (directive.startsWith(prefix)) {
					boolean isBefore = "before".equals(entry[1]);
					String expectedParent = entry[2];

					return new AugmentedDirective(prefix, directive, isBefore, expectedParent);
				}
			}
			throw new CannotParseAugmentedDirectiveException();
		}
		
		public boolean isBefore() {
			return isBefore;
		}
		
		public String getPrefix() {
			return prefix;
		}

		public String getRemaining() {
			return remaining;
		}
		
		public String getRaw() {
			return prefix + remaining;
		}

		public Object getExpectedParent() {
			return expectedParent;
		}
	}

	@Override
	public void process(OOXMLPackage pkg, Map<String, Object> rootMap) {
		InputStream f = null;
		try {
			f = new FileInputStream(new File(pkg.getDataDir(), "word/document.xml"));
			Document doc = newDocumentBuilder().parse(f);

			XPath xpath = newXPath(doc);
			NodeList nodeList = evaluateXPath(xpath, "//KMagicNode", doc);

			for (Node n : new NodeListIterAdapter(nodeList)) {
				try {
					AugmentedDirective ad = AugmentedDirective.parseDirective(n.getTextContent());
					Node runNode;
					Node parentOfPara = null;
					Node targetPara = runNode = n.getParentNode().getParentNode(); // maybe w:r

					if (!runNode.getNodeName().equals("w:r")) {
						logger.warn("runNode is not w:r({}, directive: {})", runNode.getNodeName(), ad.getRemaining());
						continue;
					}
					// find table row element following parent nodes.
					do {
						targetPara = targetPara.getParentNode();
					} while (!targetPara.getNodeName().equals(ad.getExpectedParent()));
					parentOfPara = targetPara.getParentNode();

					if (!ad.isBefore()) {
						targetPara = targetPara.getNextSibling();
					}

					// insert magic node
					parentOfPara.insertBefore(getMagicNode(doc, ad.getRemaining()), targetPara);

					// remove annotated node
					runNode.getParentNode().removeChild(runNode);
				} catch (CannotParseAugmentedDirectiveException e) {
					continue;
				}

			}

			XMLDocHelper.save(doc, new File(pkg.getDataDir(), "word/document.xml"), true);

		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			safeClose(f);
		}
	}

	private String findPrefix(String directive) {
		for (String ad : augmentedDirectives) {
			if (directive.startsWith(ad)) {
				return ad;
			}
		}
		return null;
	}

	private String unwrapAugmentedDirective(String directive) {
		for (String ad : augmentedDirectives) {
			if (directive.startsWith(ad)) {
				return directive.substring(ad.length()).trim();
			}
		}
		return directive;
	}

	private Node getMagicNode(Document doc, String content) {
		Element magicNode = doc.createElement("KMagicNode");
		magicNode.appendChild(doc.createCDATASection(content));
		return magicNode;
	}

}
