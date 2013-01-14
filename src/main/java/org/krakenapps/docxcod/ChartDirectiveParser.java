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

import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathExpressionException;

import org.apache.commons.io.FilenameUtils;
import org.krakenapps.docxcod.util.CloseableHelper;
import org.krakenapps.docxcod.util.XMLDocHelper;
import org.krakenapps.docxcod.util.XMLDocHelper.NodeListIterAdapter;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import freemarker.core.Environment;
import freemarker.template.TemplateMethodModelEx;
import freemarker.template.TemplateModelException;
import freemarker.template.utility.Collections12;

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
				throws UnsupportedDocumentException, XPathExpressionException {
			if (cells == null)
				throw new UnsupportedDocumentException("cannot obtain cell data from copied chart");

			Document chartDoc = parseXml(pkg, chartXmlPath);
			XPath xpath = newXPath(chartDoc);

			updateRef("str", chartDoc, xpath, cells);
			updateRef("num", chartDoc, xpath, cells);

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

				ArrayList<CellData> data = getDataFromRange1D(cells, formula);

				Node ptCountNode = evaluateXPath(xpath, String.format("c:%sCache/c:ptCount", pf), n).item(0);
				XMLDocHelper.setNodeAttribute(chartDoc, ptCountNode, "val", Integer.toString(data.size()));

				NodeList ptNodes = evaluateXPath(xpath, String.format("c:%sCache/c:pt", pf), n);
				if (ptNodes.getLength() < 1)
					throw new UnsupportedDocumentException("cannot update chart cache: pt node is not found");

				Node ptNode = ptNodes.item(0).cloneNode(true);
				Node cacheNode = ptNodes.item(0).getParentNode();
				for (Node c: new NodeListIterAdapter(ptNodes)) {
					c.getParentNode().removeChild(c);
				}
				int idx = 0;
				for (CellData d : data) {
					Node newPtNode = cacheNode.appendChild(ptNode.cloneNode(true));
					XMLDocHelper.setNodeAttribute(chartDoc, newPtNode, "idx", Integer.toString(idx++));
					Node vNode = evaluateXPath(xpath, "c:v", newPtNode).item(0);
					vNode.setTextContent(d.data);
				}
			}
		}

		private final Pattern cellAddressPattern = Pattern.compile("([A-Z]+)([0-9]+):([A-Z]+)([0-9]+)");
		private final Pattern cellAddressPattern2 = Pattern.compile("([A-Z]+)([0-9]+)");
		
		private ArrayList<CellData> getDataFromRange1D(ArrayList<CellData> cells, String formula) throws UnsupportedDocumentException {
			// formula example: Sheet1!$A$2:$A$4, "Sheet1!" is asserted.
			formula = formula.substring(7); // $A$2:$A$4
			formula = formula.replace("$", ""); // A2:A4
			
			if (formula.indexOf(":") < 0)
				formula = formula + ":" + formula;
			
			Matcher matcher = cellAddressPattern.matcher(formula);
			if (matcher.matches()) {
				String colId = matcher.group(1);
				int rowStart = Integer.parseInt(matcher.group(2));
				int rowEnd = Integer.parseInt(matcher.group(4)) + 1;
				
				ArrayList<CellData> data = new ArrayList<CellData>(rowEnd - rowStart);
				
				for (CellData cell: cells) {
					Matcher addrMatcher = cellAddressPattern2.matcher(cell.address);
					if (addrMatcher.matches()) {
						String cellColId = addrMatcher.group(1);
						if (!cellColId.equals(colId))
							continue;

						int cellRow = Integer.parseInt(addrMatcher.group(2));
						if (rowStart <= cellRow && cellRow < rowEnd) {
							data.add(cell);
						}	
					}
				}
				
				Collections.sort(data);
				
				return data;
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
				updateChartXML(chartXmlPath, cells);
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
