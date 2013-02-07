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

import static org.krakenapps.docxcod.Utils.readSharedStrings;
import static org.krakenapps.docxcod.util.XMLDocHelper.evaluateXPath;
import static org.krakenapps.docxcod.util.XMLDocHelper.getNodeAttribute;
import static org.krakenapps.docxcod.util.XMLDocHelper.newXPath;
import static org.krakenapps.docxcod.util.XMLDocHelper.parseXml;
import static org.krakenapps.docxcod.util.XMLDocHelper.setNodeAttribute;

import java.io.File;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathExpressionException;

import org.krakenapps.docxcod.util.XMLDocHelper;
import org.krakenapps.docxcod.util.XMLDocHelper.NodeListIterAdapter;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import freemarker.template.TemplateMethodModelEx;
import freemarker.template.TemplateModelException;

public class EmbeddedChartPreprocessor implements OOXMLProcessor {

	public static class EmbChartGetRowCntFunc implements TemplateMethodModelEx {

		public static final String name = "EmbChartGetRowCnt";
		public final AtomicInteger rowCount;

		public EmbChartGetRowCntFunc(AtomicInteger rowCount) {
			super();
			this.rowCount = rowCount;
		}

		@SuppressWarnings("rawtypes")
		@Override
		public Object exec(List arguments) throws TemplateModelException {
			return Integer.toString(rowCount.get());
		}

	}

	public static class EmbChartIncRowCntFunc implements TemplateMethodModelEx {

		public static final String name = "EmbChartIncRowCnt";
		public final AtomicInteger rowCount;

		public EmbChartIncRowCntFunc(AtomicInteger rowCount) {
			super();
			this.rowCount = rowCount;
		}

		@SuppressWarnings("rawtypes")
		@Override
		public Object exec(List arguments) throws TemplateModelException {
			// TODO increase row count
			int value = rowCount.incrementAndGet();
			return Integer.toString(value);
		}

	}

	public static class EmbChartCellRefFunc implements TemplateMethodModelEx {

		public static final String name = "EmbChartCellRef";

		@SuppressWarnings("rawtypes")
		@Override
		public Object exec(List arguments) throws TemplateModelException {
			// TODO generate A1 type cell reference while increasing row ref.
			// arg1 : row count
			// arg2 : cell count
			int rowCount = Integer.parseInt(arguments.get(0).toString());
			char cellIdx = (char) ('A' + (-1 + Integer.parseInt(arguments.get(1).toString())));
			return String.format("%s%d", cellIdx, rowCount);
		}

	}

	private static final Pattern cellAddressPattern = Pattern.compile("[A-Z]+[0-9]+:([A-Z]+)([0-9]+)");
	private Logger logger = LoggerFactory.getLogger(this.getClass().getName());

	@Override
	public void process(OOXMLPackage pkg, Map<String, Object> rootMap) {
		try {
			if (rootMap != null) {
				AtomicInteger atomicInteger = new AtomicInteger(1);
				rootMap.put(EmbChartGetRowCntFunc.name, new EmbChartGetRowCntFunc(atomicInteger));
				rootMap.put(EmbChartIncRowCntFunc.name, new EmbChartIncRowCntFunc(atomicInteger));
				rootMap.put(EmbChartCellRefFunc.name, new EmbChartCellRefFunc());
			}

			Document sheet1Doc = parseXml(pkg, "xl/worksheets/sheet1.xml");
			Document sharedStringsDoc = parseXml(pkg, "xl/sharedStrings.xml");

			List<String> sharedStrings = readSharedStrings(sharedStringsDoc);
			LoopDescriptor loopDescriptor = readDocxcodAnnotation(sheet1Doc, sharedStrings);

			modifySheet1Xml(sheet1Doc, loopDescriptor);

			XMLDocHelper.save(sheet1Doc, new File(pkg.getDataDir(), "xl/worksheets/sheet1.xml"), true);

			logger.info("sheet modification completed");

		} catch (Exception e) {
			e.printStackTrace();
		} finally {
		}
	}

	private void modifySheet1Xml(Document sheet1Doc, LoopDescriptor loopDesc) throws XPathExpressionException,
			UnsupportedDocumentException {
		XPath xpath = newXPath(sheet1Doc);

		NodeList nodeList = evaluateXPath(xpath, "//DEF:sheetData/DEF:row[@r>1]", sheet1Doc); // XXX
		Node rowTemplate = nodeList.item(0).cloneNode(true);
		Node parentNode = nodeList.item(0).getParentNode();

		// add new rows
		parentNode.insertBefore(rowTemplate, nodeList.item(nodeList.getLength() - 1).getNextSibling());
		
		// remove existing rows
		for (Node n : new NodeListIterAdapter(nodeList)) {
			n.getParentNode().removeChild(n);
		}

		Node dimNode = evaluateXPath(xpath, "//DEF:dimension", sheet1Doc).item(0);
		String dim = getNodeAttribute(dimNode, "ref");
		Matcher matcher = cellAddressPattern.matcher(dim);
		int rightEnd = 16;
		if (matcher.matches()) {
			rightEnd = matcher.group(1).charAt(0) - 'A' + 1;
		}

		setNodeAttribute(sheet1Doc, rowTemplate, "r", String.format("${%s()}", EmbChartIncRowCntFunc.name));
		setNodeAttribute(sheet1Doc, rowTemplate, "spans", String.format("1:%d", rightEnd));

		// modify cell definitions
		NodeList cells = evaluateXPath(xpath, "DEF:c", rowTemplate);
		int cellCount = 0;
		for (Node n : new NodeListIterAdapter(cells)) {
			cellCount += 1;

			// remove cells out of span.
			if (cellCount > loopDesc.itemDescriptors.size()) {
				n.getParentNode().removeChild(n);
				continue;
			}

			ItemDescriptor itemDesc = loopDesc.itemDescriptors.get(cellCount - 1);

			setNodeAttribute(sheet1Doc, n, "r",
					String.format("${%s(%s(), %d)}", EmbChartCellRefFunc.name, EmbChartGetRowCntFunc.name, cellCount));
			// modify content of cell
			if ("s".equals(getNodeAttribute(n, "t"))) {
				// if cell type is string, use inlineStr
				modifySharedStringCell(sheet1Doc, n, itemDesc);
			} else {
				modifyValueCell(n, itemDesc);
			}
		}
		
		// fill empty cell 
		if (rightEnd > loopDesc.itemDescriptors.size()) {
			for (int cc = loopDesc.itemDescriptors.size() + 1; cc <= rightEnd; ++cc) {
				rowTemplate.appendChild(createCellNode(sheet1Doc, cc));
			}
		}
		
		// append magic node before and after rowTemplate
		rowTemplate.getParentNode().insertBefore(getMagicNode(sheet1Doc, loopDesc.loopCoordinator), rowTemplate);
		rowTemplate.getParentNode().insertBefore(getMagicNode(sheet1Doc, "/#list"), rowTemplate.getNextSibling());

	}

	private Node createCellNode(Document sheet1Doc, int cellCount) {
		Element cellNode = sheet1Doc.createElement("c");
		setNodeAttribute(sheet1Doc, cellNode, "r",
				String.format("${%s(%s(), %d)}", EmbChartCellRefFunc.name, EmbChartGetRowCntFunc.name, cellCount));
		setNodeAttribute(sheet1Doc, cellNode, "s", "1");
		
		return cellNode;
	}

	private Node getMagicNode(Document doc, String content) {
		Element magicNode = doc.createElement("KMagicNode");
		magicNode.appendChild(doc.createCDATASection(content));
		return magicNode;
	}

	private void modifyValueCell(Node n, ItemDescriptor itemDesc) throws UnsupportedDocumentException {
		Node vElem = null;
		for (Node c : new NodeListIterAdapter(n.getChildNodes())) {
			if ("v".equals(c.getNodeName())) {
				vElem = c;
				break;
			}
		}
		if (vElem == null)
			throw new UnsupportedDocumentException("cannot modify cell content: no v elem in c elem");
		vElem.setTextContent(itemDesc.sharedString);
	}

	private void modifySharedStringCell(Document sheet1Doc, Node n, ItemDescriptor itemDesc) {
		setNodeAttribute(sheet1Doc, n, "t", "inlineStr");
		NodeList childNodes = n.getChildNodes();
		while(childNodes.getLength() > 0) {
			Node c = childNodes.item(0);
			c.getParentNode().removeChild(c);
		}
		n.appendChild(createInlineStrElement(sheet1Doc, itemDesc.sharedString));
	}

	private Node createInlineStrElement(Document sheet1Doc, String sharedString) {
		Element isElem = sheet1Doc.createElement("is");
		Element tElem = sheet1Doc.createElement("t");
		tElem.setTextContent(sharedString);
		isElem.appendChild(tElem);
		return isElem;
	}

	public static class ItemDescriptor {

		public final String refAddr;
		public final String sharedString;

		public ItemDescriptor(String refAddr, String sharedString) {
			this.refAddr = refAddr;
			this.sharedString = sharedString;
		}

		@Override
		public String toString() {
			return "[refAddr=" + refAddr + ", sharedString=" + sharedString + "]";
		}

	}

	private static class LoopDescriptor {
		public final String loopCoordinator;
		public final ArrayList<ItemDescriptor> itemDescriptors;

		public LoopDescriptor(String loopCoordinator, ArrayList<ItemDescriptor> itemDescriptors) {
			this.loopCoordinator = loopCoordinator;
			this.itemDescriptors = itemDescriptors;
		}

		@Override
		public String toString() {
			return "[loopCoordinator=" + loopCoordinator + ", itemDescriptors="
					+ Arrays.toString(itemDescriptors.toArray()) + "]";
		}

	}

	private LoopDescriptor readDocxcodAnnotation(Document sheet1Doc, List<String> sst) throws XPathExpressionException,
			UnsupportedDocumentException {

		XPath xpath = newXPath(sheet1Doc);

		NodeList values = null;
		String loopCoordinator = null;
		ArrayList<ItemDescriptor> itemDescriptors = new ArrayList<ItemDescriptor>();

		values = evaluateXPath(xpath, "//*[name()='row' and DEF:c][last()-1]/DEF:c[1]/DEF:v", sheet1Doc);
		if (values.getLength() < 1)
			throw new UnsupportedDocumentException("cannot obtain loop coordinator");
		else {
			int i = Integer.parseInt(values.item(0).getTextContent());
			loopCoordinator = getSharedString(sst, i);
		}

		values = evaluateXPath(xpath, "//*[name()='row' and DEF:c][last()]/DEF:c", sheet1Doc);
		if (values.getLength() < 1)
			throw new UnsupportedDocumentException("cannot obtain item descriptor");
		for (Node n : new NodeListIterAdapter(values)) {
			itemDescriptors.add(createItemDescriptor(xpath, sst, n));
		}

		return new LoopDescriptor(loopCoordinator, itemDescriptors);

	}

	private ItemDescriptor createItemDescriptor(XPath xpath, List<String> sst, Node n) throws XPathExpressionException,
			UnsupportedDocumentException {
		/*
		 * <c r="A8" s="2" t="s"> <v>6</v> </c>
		 */

		String refAddr = n.getAttributes().getNamedItem("r").getNodeValue();
		Node v = evaluateXPath(xpath, "DEF:v", n).item(0);
		int ssIdx = Integer.parseInt(v.getTextContent());

		return new ItemDescriptor(refAddr, getSharedString(sst, ssIdx));
	}

	private String getSharedString(List<String> sst, int i) throws UnsupportedDocumentException {
		if (i < sst.size())
			return sst.get(i);
		else
			throw new UnsupportedDocumentException("shared string map doesn't contains item #" + i);
	}

}
