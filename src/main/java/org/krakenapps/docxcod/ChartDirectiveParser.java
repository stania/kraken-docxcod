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
import static org.krakenapps.docxcod.util.XMLDocHelper.newDocumentBuilder;
import static org.krakenapps.docxcod.util.XMLDocHelper.newXPath;
import static org.krakenapps.docxcod.util.XMLDocHelper.parseXml;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactoryConfigurationError;
import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathExpressionException;

import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang.StringUtils;
import org.krakenapps.docxcod.util.CloseableHelper;
import org.krakenapps.docxcod.util.XMLDocHelper;
import org.krakenapps.docxcod.util.XMLDocHelper.NodeListIterAdapter;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.w3c.dom.DOMException;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import com.google.common.base.Predicate;
import com.google.common.collect.Collections2;

import freemarker.core.Environment;
import freemarker.template.TemplateMethodModelEx;
import freemarker.template.TemplateModelException;

public class ChartDirectiveParser implements OOXMLProcessor {
	// returns unique incremental integer value for consecutive same originalRid
	public class ChartUidFunction implements TemplateMethodModelEx {
		public static final String functionName = "chartUid";
		private String rid = null;
		private int count = 0;

		@Override
		public Object exec(@SuppressWarnings("rawtypes") List arguments) throws TemplateModelException {
			String origRid = arguments.get(0).toString();
			if (!origRid.equals(rid)) {
				rid = origRid;
				count = 0;
			}
			return Integer.toString(count++);
		}
	}

	private static final String CHART_XML_CONTENTTYPE = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
	private static final String DOCXCOD_CHART_XML_EXT = "docxcod_chart_xml";
	private static final String CONTENT_TYPES_XML = "[Content_Types].xml";

	public class ChartResFunction implements TemplateMethodModelEx {
		public static final String functionName = "ridHelper";
		public int count = 0;
		private boolean contentTypeAppended = false;
		private final OOXMLPackage pkg;

		public ChartResFunction(OOXMLPackage pkg) {
			this.pkg = pkg;
		}

		@Override
		public Object exec(@SuppressWarnings("rawtypes") List arguments) throws TemplateModelException {
			try {
				String originalRid = arguments.get(0).toString();
				String chartUid = arguments.get(1).toString();
				String rid = originalRid + "_" + chartUid;

				Environment ce = Environment.getCurrentEnvironment();
				@SuppressWarnings("unchecked")
				Set<String> knownVariable = ce.getKnownVariableNames();
				HashMap<String, Object> localRoot = new HashMap<String, Object>();
				for (String k : knownVariable) {
					localRoot.put(k, ce.getVariable(k));
				}
				// add ".kraken_chart_xml" to [Content_Types].xml
				appendContentType(pkg);
				// create copied chart(xml, xlsx)
				createCopiedChart(pkg, originalRid, chartUid, localRoot);

				// must return text content of attr "r:id".
				return rid;
			} catch (Exception e) {
				logger.warn("Unhandled exception", e);
			}
			// if this dd
			return null;
		}

		private void updateChartXML(String chartXmlPath, ArrayList<CellData> cells)
				throws UnsupportedDocumentException, XPathExpressionException, DOMException,
				TransformerFactoryConfigurationError, TransformerException {
			if (cells == null)
				throw new UnsupportedDocumentException("cannot obtain cell data from copied chart");

			Document chartDoc = parseXml(pkg, chartXmlPath);
			XPath xpath = newXPath(chartDoc);

			updateRef("str", chartDoc, xpath, cells);
			updateRef("num", chartDoc, xpath, cells);

			XMLDocHelper.save(chartDoc, new File(pkg.getDataDir(), chartXmlPath), true);
		}

		private void updateRef(String pf, Document chartDoc, XPath xpath, ArrayList<CellData> cells)
				throws XPathExpressionException, UnsupportedDocumentException {

			NodeList nodeList = evaluateXPath(xpath, String.format("//c:%sRef", pf), chartDoc);
			for (Node n : new NodeListIterAdapter(nodeList)) {
				Node fNode = evaluateXPath(xpath, "c:f", n).item(0);
				if (fNode == null)
					throw new UnsupportedDocumentException("cannot update chart cache: no formula node in ref elem");

				String formula = fNode.getTextContent();
				if (!formula.startsWith("Sheet1!"))
					throw new UnsupportedDocumentException("cannot update chart cache: not referencing Sheet1");

				ArrayList<CellData> data = getColumnDataFromRange1D(cells, formula);

				fNode.setTextContent(getModifiedFormula(formula, data.size()));

				Node ptCountNode = evaluateXPath(xpath, String.format("c:%sCache/c:ptCount", pf), n).item(0);
				XMLDocHelper.setNodeAttribute(chartDoc, ptCountNode, "val", Integer.toString(data.size()));

				NodeList ptNodes = evaluateXPath(xpath, String.format("c:%sCache/c:pt", pf), n);
				if (ptNodes.getLength() < 1)
					throw new UnsupportedDocumentException("cannot update chart cache: pt node is not found");

				Node ptNode = ptNodes.item(0).cloneNode(true);
				Node cacheNode = ptNodes.item(0).getParentNode();
				for (Node c : new NodeListIterAdapter(ptNodes)) {
					c.getParentNode().removeChild(c);
				}
				int idx = 0;
				if (data.isEmpty()) {
					// if chart contains no data, add 1 cell which contains 0
					Node newPtNode = cacheNode.appendChild(ptNode.cloneNode(true));
					XMLDocHelper.setNodeAttribute(chartDoc, newPtNode, "idx", Integer.toString(idx++));
					Node vNode = evaluateXPath(xpath, "c:v", newPtNode).item(0);
					vNode.setTextContent("0");
				} else {
					// normal case
					for (CellData d : data) {
						Node newPtNode = cacheNode.appendChild(ptNode.cloneNode(true));
						XMLDocHelper.setNodeAttribute(chartDoc, newPtNode, "idx", Integer.toString(idx++));
						Node vNode = evaluateXPath(xpath, "c:v", newPtNode).item(0);
						vNode.setTextContent(d.data);
					}
				}
			}
		}

		private String getModifiedFormula(String formula, int size) {
			/*
			 * ex input of formula: Sheet1!$B$2:$B$4 ex input of size: 0
			 */
			if (size == 0)
				size = 1; // fix errorneous data

			Matcher matcher = cellAbsAddressPattern.matcher(formula);
			if (matcher.matches()) {
				int startRow = Integer.parseInt(matcher.group(2));
				formula = formula.substring(0, matcher.start(4)) + Integer.toString(startRow + size - 1)
						+ formula.substring(matcher.end(4));
			}
			return formula;
		}

		private final Pattern cellAddressPattern = Pattern.compile("([A-Z]+)([0-9]+):([A-Z]+)([0-9]+)");
		private final Pattern cellAbsAddressPattern = Pattern
				.compile("Sheet1!\\$([A-Z]+)\\$([0-9]+):\\$([A-Z]+)\\$([0-9]+)");
		private final Pattern cellAddressPattern2 = Pattern.compile("([A-Z]+)([0-9]+)");

		private ArrayList<CellData> getColumnDataFromRange1D(ArrayList<CellData> cells, String formula)
				throws UnsupportedDocumentException {
			// formula example: Sheet1!$A$2:$A$4, "Sheet1!" is asserted.
			formula = formula.substring(7); // $A$2:$A$4
			formula = formula.replace("$", ""); // A2:A4

			if (formula.indexOf(":") < 0)
				formula = formula + ":" + formula;

			Matcher matcher = cellAddressPattern.matcher(formula);
			if (matcher.matches()) {
				String colId = matcher.group(1);
				int rowStart = Integer.parseInt(matcher.group(2));
				int rowEnd = Integer.parseInt(matcher.group(4));

				ArrayList<CellData> data = new ArrayList<CellData>();

				if (rowStart == rowEnd) {
					String targetAddr = String.format("%s%d", colId, rowStart);
					for (CellData cell : cells) {
						if (targetAddr.equals(cell.address)) {
							data.add(cell);
							break;
						}
					}

					return data;
				} else {

					for (CellData cell : cells) {
						Matcher addrMatcher = cellAddressPattern2.matcher(cell.address);
						if (addrMatcher.matches()) {
							String cellColId = addrMatcher.group(1);
							if (!cellColId.equals(colId))
								continue;

							int cellRow = Integer.parseInt(addrMatcher.group(2));
							if (rowStart <= cellRow) {
								data.add(cell);
							}
						}
					}

					Collections.sort(data);

					// filter cells containing data continuously
					data = new ArrayList<CellData>(Collections2.filter(data, new Predicate<CellData>() {
						int prevRow = 0;

						@Override
						public boolean apply(CellData input) {
							int rowIdx = Integer.parseInt(input.address.substring(StringUtils.indexOfAny(input.address,
									"0123456789")));
							if (prevRow == 0) {
								prevRow = rowIdx;
								return true;
							} else {
								if (rowIdx - prevRow == 1) {
									prevRow = rowIdx;
									return true;
								} else {
									return false;
								}
							}
						}
					}));

					return data;
				}
			} else {
				throw new UnsupportedDocumentException("internal error: cannot match cell formula");
			}
		}

		private void appendContentType(OOXMLPackage pkg) {
			InputStream f = null;
			try {
				f = new FileInputStream(new File(pkg.getDataDir(), CONTENT_TYPES_XML));
				Document doc = newDocumentBuilder().parse(f);

				XPath xpath = newXPath(doc);
				NodeList nodeList = evaluateXPath(xpath, "//:Default[@Extension='" + DOCXCOD_CHART_XML_EXT + "']", doc);
				if (!contentTypeAppended && nodeList.getLength() == 0) {
					Element newChild = doc.createElement("Default");
					newChild.setAttribute("ContentType", CHART_XML_CONTENTTYPE);
					newChild.setAttribute("Extension", DOCXCOD_CHART_XML_EXT);
					doc.getFirstChild().appendChild(newChild);

					XMLDocHelper.save(doc, new File(pkg.getDataDir(), CONTENT_TYPES_XML), true);

					this.contentTypeAppended = true;
				}

			} catch (Exception e) {
				e.printStackTrace();
			} finally {
				CloseableHelper.safeClose(f);
			}
		}

		private void createCopiedChart(OOXMLPackage pkg, String originalRid, String chartUid,
				HashMap<String, Object> localRoot) throws Exception {
			// append new chart xml entry in ooxml relationship file.
			String chartXmlPath = appendToRels(pkg, "word/_rels/document.xml.rels", originalRid, chartUid);

			if (chartXmlPath != null) {
				// translate chart xml path to more absolute path
				// (ex: word/chart/chart1_0.docxcod_chart_xml)
				chartXmlPath = FilenameUtils.concat("word", chartXmlPath);
				logger.info("new chart xml filename: {}", chartXmlPath);

				// copy chart xml to chartXmlPath and change reference to xlsx
				String embeddedXlsxPath = createChartFromExisting(pkg, chartXmlPath, chartUid);

				// create new xlsx for chart in chartXmlPath
				ArrayList<CellData> cells = createNewEmbeddedXlsx(pkg, embeddedXlsxPath, chartUid, localRoot);

				// update strRef/numRef in chart xml with embeddedXlsx
				updateChartXML(makeNewChartFilename(chartXmlPath, chartUid), cells);
			} else {
				throw new Exception("cannot append rels entry in document.xml.rels");
			}
		}

		private String createChartFromExisting(OOXMLPackage pkg, String chartXmlPath, String chartUid) {
			InputStream f = null;
			try {
				// create copied chart xml
				String embeddedXlsxRid = null;
				String embeddedXlsxFile = null;

				f = new FileInputStream(new File(pkg.getDataDir(), chartXmlPath));
				Document doc = newDocumentBuilder().parse(f);

				XPath xpath = newXPath(doc);
				NodeList nodeList = evaluateXPath(xpath, "//c:externalData", doc);
				if (nodeList.getLength() != 0) {
					Node n = nodeList.item(0);
					Node attrRid = n.getAttributes().getNamedItem("r:id");
					if (attrRid == null)
						throw new IllegalStateException(String.format("no Target externalData in %s", chartXmlPath));
					embeddedXlsxRid = attrRid.getTextContent();
				}

				XMLDocHelper.save(doc, new File(pkg.getDataDir(), makeNewChartFilename(chartXmlPath, chartUid)), true);

				f.close();

				// open newly created xml and modify relationship to xlsx
				f = new FileInputStream(new File(pkg.getDataDir(), makeRelsPath(chartXmlPath)));
				doc = newDocumentBuilder().parse(f);
				xpath = newXPath(doc);
				if (embeddedXlsxRid != null) {
					NodeList rsNodes = evaluateXPath(xpath, "//:Relationship[@Id='" + embeddedXlsxRid + "']", doc);
					if (rsNodes.getLength() != 0) {
						Node n = rsNodes.item(0);
						embeddedXlsxFile = FilenameUtils.concat("word/charts", n.getAttributes().getNamedItem("Target")
								.getTextContent());
						String relTarget = makeXlsxRelTarget(chartXmlPath,
								makeNewXlsxFilename(embeddedXlsxFile, chartUid));
						n.getAttributes().getNamedItem("Target")
								.setTextContent(FilenameUtils.separatorsToUnix(relTarget));
					}
				}
				XMLDocHelper.save(doc,
						new File(pkg.getDataDir(), makeRelsPath(makeNewChartFilename(chartXmlPath, chartUid))), true);

				return embeddedXlsxFile;

			} catch (Exception e) {
				e.printStackTrace();
			} finally {
				CloseableHelper.safeClose(f);
			}

			return null;
		}

		private String makeRelsPath(String chartXmlPath) {
			String name = FilenameUtils.getName(chartXmlPath);
			String path = FilenameUtils.getPath(chartXmlPath);

			return FilenameUtils.concat(FilenameUtils.concat(path, "_rels"), name + ".rels");
		}

		private String makeXlsxRelTarget(String chartXmlPath, String xlsxFilename) {
			// TODO: use reliable relative path logic
			return "../embeddings/" + FilenameUtils.getName(xlsxFilename);
		}

		private String makeNewChartFilename(String chartXmlPath, String chartUid) {
			// trim right ".xml"
			String s = chartXmlPath.substring(0, chartXmlPath.length() - 4);
			s += "_" + chartUid + "." + DOCXCOD_CHART_XML_EXT;
			return s;
		}

		private String appendToRels(OOXMLPackage pkg, String relPath, String originalRid, String chartUid) {
			InputStream f = null;
			try {
				f = new FileInputStream(new File(pkg.getDataDir(), relPath));
				Document doc = newDocumentBuilder().parse(f);

				XPath xpath = newXPath(doc);
				NodeList nodeList = evaluateXPath(xpath, "//:Relationship[@Id='" + originalRid + "']", doc);
				String chartXmlPath = null;
				if (nodeList.getLength() != 0) {
					Node n = nodeList.item(0);

					Node attrTarget = n.getAttributes().getNamedItem("Target");
					if (attrTarget == null)
						throw new IllegalStateException(String.format("no Target attribute in %s with rid %s", relPath,
								originalRid));
					chartXmlPath = attrTarget.getTextContent();

					Node attrType = n.getAttributes().getNamedItem("Type");
					if (attrType == null)
						throw new IllegalStateException(String.format("no Type attribute in %s with rid %s", relPath,
								originalRid));
					String type = attrType.getTextContent();

					Element newChild = doc.createElement("Relationship");

					newChild.setAttribute("Id", originalRid + "_" + chartUid);
					newChild.setAttribute("Target",
							FilenameUtils.separatorsToUnix(makeNewChartFilename(chartXmlPath, chartUid)));
					newChild.setAttribute("Type", type);
					doc.getFirstChild().appendChild(newChild);
				} else {
					return null;
				}

				XMLDocHelper.save(doc, new File(pkg.getDataDir(), relPath), true);

				return chartXmlPath;

			} catch (Exception e) {
				e.printStackTrace();
			} finally {
				CloseableHelper.safeClose(f);
			}

			return null;
		}

		private String makeNewXlsxFilename(String embeddedXlsxFile, String chartUid) {
			String path = FilenameUtils.getPath(embeddedXlsxFile);
			String basename = FilenameUtils.getBaseName(embeddedXlsxFile);
			String ext = FilenameUtils.getExtension(embeddedXlsxFile);

			return path + basename + "_" + chartUid + "." + ext;
		}

		private ArrayList<CellData> createNewEmbeddedXlsx(OOXMLPackage pkg, String embeddedXlsxFile, String chartUid,
				Map<String, Object> localRoot) {
			OOXMLPackage xlsx = new OOXMLPackage();
			FileInputStream is = null;
			FileOutputStream os = null;
			try {
				is = new FileInputStream(new File(pkg.getDataDir(), embeddedXlsxFile));
				File targetDir = new File(".test/xlsxSample");
				targetDir.mkdirs();
				xlsx.load(is, targetDir);

				ArrayList<CellData> cells = new ArrayList<CellData>();

				List<OOXMLProcessor> processors = new ArrayList<OOXMLProcessor>();
				processors.add(new EmbeddedChartPreprocessor());
				processors.add(new MagicNodeUnwrapper("xl/worksheets/sheet1.xml"));
				processors.add(new FreeMarkerRunner("xl/worksheets/sheet1.xml"));
				processors.add(new EmbeddedChartPostprocessor(cells));

				xlsx.apply(processors, localRoot);

				os = new FileOutputStream(new File(pkg.getDataDir(), makeNewXlsxFilename(embeddedXlsxFile, chartUid)));
				xlsx.save(os);

				return cells;

			} catch (FileNotFoundException e) {
				e.printStackTrace();
			} catch (IOException e) {
				e.printStackTrace();
			} finally {
				CloseableHelper.safeClose(is);
				CloseableHelper.safeClose(os);
			}

			return null;
		}

		public String getName() {
			return functionName;
		}
	}

	private Logger logger = LoggerFactory.getLogger(getClass().getName());

	@Override
	public void process(OOXMLPackage pkg, Map<String, Object> rootMap) {
		if (rootMap != null) {
			rootMap.put(ChartResFunction.functionName, new ChartResFunction(pkg));
			rootMap.put(ChartUidFunction.functionName, new ChartUidFunction());
		}

		InputStream f = null;
		try {
			f = new FileInputStream(new File(pkg.getDataDir(), "word/document.xml"));
			Document doc = newDocumentBuilder().parse(f);

			XPath xpath = newXPath(doc);

			NodeList nodeList = evaluateXPath(xpath, "//c:chart", doc);

			for (Node n : new NodeListIterAdapter(nodeList)) {
				InsertChartHelperMagicNode(doc, n);
			}

			XMLDocHelper.save(doc, new File(pkg.getDataDir(), "word/document.xml"), true);

		} catch (XPathExpressionException e) {
			// "접두부는 이름 공간으로 분석되어야 합니다: c"
			logger.trace("maybe no chart element in this document. pass.");
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			CloseableHelper.safeClose(f);
		}
	}

	private enum AsttpPos {
		BEFORE,
		AFTER
	};

	private void InsertChartHelperMagicNode(Document doc, Node chartNode) {
		Node attrRid = chartNode.getAttributes().getNamedItem("r:id");
		String originalRid = attrRid.getTextContent();
		logger.info("chart rid: {}", originalRid);
		// ChartUidFunction returns unique incremental integer value for
		// consecutive same originalRid
		appendSiblingToTheParent(
				chartNode,
				"w:drawing",
				AsttpPos.BEFORE,
				getMagicNode(doc,
						String.format("#assign curChartUid=%s(\'%s\')", ChartUidFunction.functionName, originalRid)));
		attrRid.setTextContent(String.format("${%s(\'%s\', curChartUid)}", ChartResFunction.functionName, originalRid));
	}

	private boolean appendSiblingToTheParent(Node currentNode, String parentNodeName, AsttpPos posOfSibling,
			Node magicNode) {
		Node p = currentNode;
		for (;;) {
			Node t = p.getParentNode();
			if (t == null)
				return false;
			p = t;
			if (p.getNodeName().equals(parentNodeName))
				break;
		}
		if (p.getParentNode() == null)
			return false;

		switch (posOfSibling) {
		case BEFORE:
			p.getParentNode().insertBefore(magicNode, p);
			break;
		case AFTER:
			p.getParentNode().insertBefore(magicNode, p.getNextSibling());
			break;
		}
		return true;
	}

	private Node getMagicNode(Document doc, String content) {
		Element magicNode = doc.createElement("KMagicNode");
		magicNode.appendChild(doc.createCDATASection(content));
		return magicNode;
	}

}
