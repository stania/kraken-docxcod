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
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactoryConfigurationError;
import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathExpressionException;

import org.krakenapps.docxcod.util.XMLDocHelper;
import org.krakenapps.docxcod.util.XMLDocHelper.NodeListIterAdapter;
import org.w3c.dom.DOMException;
import org.w3c.dom.Document;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

public class EmbeddedChartPostprocessor implements OOXMLProcessor {

	private final ArrayList<CellData> cells;

	public EmbeddedChartPostprocessor(ArrayList<CellData> cells) {
		this.cells = cells;
	}

	@Override
	public void process(OOXMLPackage pkg, Map<String, Object> rootMap) {
		
		try {
			// read sheet1.xml
			String sheetXmlFile = "xl/worksheets/sheet1.xml";

			Document sheet1Doc = parseXml(pkg, sheetXmlFile);
			Document sharedStringsDoc = parseXml(pkg, "xl/sharedStrings.xml");

			List<String> sharedStrings = readSharedStrings(sharedStringsDoc);

			readSheetData(sheet1Doc, sharedStrings, cells);

			String tableRange = getTableRange(cells);

			// update xl/tables/table1.xml
			updateTable1Xml(pkg, tableRange);

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	private void updateTable1Xml(OOXMLPackage pkg, String tableRange) throws XPathExpressionException, DOMException,
			TransformerFactoryConfigurationError, TransformerException {
		Document table1Doc = parseXml(pkg, "xl/tables/table1.xml");
		Node tableNode = evaluateXPath(table1Doc, "/DEF:table").item(0);
		setNodeAttribute(table1Doc, tableNode, "ref", tableRange);
		XMLDocHelper.save(table1Doc, new File(pkg.getDataDir(), "xl/tables/table1.xml"), false);
	}

	private String getTableRange(List<CellData> cells) {
		int row = 0;
		int col = 0;

		HashMap<String, CellData> cellAddrs = new HashMap<String, CellData>();
		for (CellData c : cells) {
			cellAddrs.put(c.address, c);
		}

		while (doesCellContainsData(row + 1, col, cellAddrs)) {
			row++;
		}
		while (doesCellContainsData(row, col + 1, cellAddrs)) {
			col++;
		}

		return String.format("A1:%s", getCellAddr(row, col));
	}

	private boolean doesCellContainsData(int row, int col, HashMap<String, CellData> cellAddrs) {
		String cellAddr = getCellAddr(row, col);
		return cellAddrs.containsKey(cellAddr) && !cellAddrs.get(cellAddr).data.isEmpty();
	}

	private String getCellAddr(int row, int col) {
		char cellIdx = (char) ('A' + col);
		return cellIdx + Integer.toString(row + 1);
	}

	private void readSheetData(Document sheet1Doc, List<String> sharedStrings, ArrayList<CellData> cells)
			throws XPathExpressionException {
		XPath xpath = newXPath(sheet1Doc);
		NodeList cellNodes = evaluateXPath(xpath, "//DEF:c", sheet1Doc);
		cells.ensureCapacity(cellNodes.getLength());

		for (Node n : new NodeListIterAdapter(cellNodes)) {
			String address;
			String data;

			address = getNodeAttribute(n, "r");
			String typeAttr = getNodeAttribute(n, "t");
			if ("s".equals(typeAttr)) {
				Node item = evaluateXPath(xpath, "DEF:v", n).item(0);
				if (item != null) {
					Integer sstIdx = Integer.parseInt(item.getTextContent());
					if (sharedStrings.size() < sstIdx)
						data = "";
					else
						data = sharedStrings.get(sstIdx);
				} else {
					data = "";
				}
			} else if ("inlineStr".equals(typeAttr)) {
				data = evaluateXPath(xpath, "DEF:is/DEF:t", n).item(0).getTextContent();
			} else {
				Node item = evaluateXPath(xpath, "DEF:v", n).item(0);
				if (item != null)
					data = item.getTextContent();
				else 
					data = "";
			}

			cells.add(new CellData(address, data));
		}
	}

}
