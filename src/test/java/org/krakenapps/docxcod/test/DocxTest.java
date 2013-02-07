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

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

import org.json.JSONException;
import org.json.JSONObject;
import org.json.JSONTokener;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import org.krakenapps.docxcod.AugmentedDirectiveProcessor;
import org.krakenapps.docxcod.ChartDirectiveParser;
import org.krakenapps.docxcod.Directive;
import org.krakenapps.docxcod.DirectiveExtractor;
import org.krakenapps.docxcod.FreeMarkerRunner;
import org.krakenapps.docxcod.JsonHelper;
import org.krakenapps.docxcod.MagicNodeUnwrapper;
import org.krakenapps.docxcod.MergeFieldParser;
import org.krakenapps.docxcod.OOXMLPackage;
import org.krakenapps.docxcod.OOXMLProcessor;
import org.krakenapps.docxcod.util.ZipHelper;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.w3c.dom.Node;

public class DocxTest {
	private Logger logger = LoggerFactory.getLogger(getClass().getName());

	private TearDownHelper tearDownHelper = new TearDownHelper();

	@Before
	public void setUp() throws Exception {
	}

	@After
	public void tearDown() throws Exception {
		tearDownHelper.tearDown();
	}
	
	@Test
	public void fieldTest() throws IOException {
		File targetDir = new File(".test/fieldTest");
		targetDir.mkdirs();
		// tearDownHelper.add(targetDir);

		OOXMLPackage docx = new OOXMLPackage();
		docx.load(getClass().getResourceAsStream("/fieldTest.docx"), targetDir);

		DirectiveExtractor directiveExtractor = new DirectiveExtractor();

		docx.apply(directiveExtractor, null);

		String[] expected = new String[] {
				"@before-row#list data as a asdfasdfasdsfd",
				"asdsfdadsfadsfsdfa",
				"aslasdfasdfasdfddkfj",
				"asdfasdfasdf",
				"@before-row #list .vars[\"stania\"] as a",
				"asdfadfasdf",
				"ahsdfl;aksjdflksdjf",
				"#list asdf as qwer",
				"@before-row#list data as row",
				"@before-row#list data as row" // there're hidden instrText
		};

		int cnt = 0;
		for (Directive dir : directiveExtractor.getDirectives()) {
			@SuppressWarnings("unused")
			Node n = dir.getPosition();
			String dirStr = dir.getDirectiveString();
			logger.debug("extracted: " + dirStr);

			assertTrue(dirStr.equals(expected[cnt++]));
		}
	}

	@Test
	public void totalTest() throws IOException, JSONException {
		File targetDir = new File(".test/_totalTest");
		targetDir.mkdirs();
		// tearDownHelper.add(targetDir);

		OOXMLPackage docx = new OOXMLPackage();
		docx.load(getClass().getResourceAsStream("/totalTest.docx"), targetDir);

		InputStreamReader inputReader = new InputStreamReader(getClass().getResourceAsStream("/totalTest.in"));
		JSONTokener tokener = new JSONTokener(inputReader);
		Map<String, Object> rootMap = JsonHelper.parse((JSONObject) tokener.nextValue());

		List<OOXMLProcessor> processors = new ArrayList<OOXMLProcessor>();
		docx.apply(new MergeFieldParser(), rootMap);
		docx.apply(new AugmentedDirectiveProcessor(), rootMap);
		docx.apply(new ChartDirectiveParser(), rootMap);
		docx.apply(new MagicNodeUnwrapper("word/document.xml"), rootMap);
		docx.apply(new FreeMarkerRunner("word/document.xml"), rootMap);

		File saveFile = new File(".test/totalTest-save.docx");
		docx.save(new FileOutputStream(saveFile));
		// tearDownHelper.add(saveFile);
		
	}


	@Test
	public void chartTest() throws IOException, JSONException {
		File targetDir = new File(".test/_chartTest");
		targetDir.mkdirs();
		// tearDownHelper.add(targetDir);

		OOXMLPackage docx = new OOXMLPackage();
		docx.load(getClass().getResourceAsStream("/chartTest.docx"), targetDir);

		InputStreamReader inputReader = new InputStreamReader(getClass().getResourceAsStream("/nestedListTest.in"));
		JSONTokener tokener = new JSONTokener(inputReader);
		Map<String, Object> rootMap = JsonHelper.parse((JSONObject) tokener.nextValue());

		List<OOXMLProcessor> processors = new ArrayList<OOXMLProcessor>();
		processors.add(new MergeFieldParser());
		processors.add(new AugmentedDirectiveProcessor());
		processors.add(new ChartDirectiveParser());
		processors.add(new MagicNodeUnwrapper("word/document.xml"));
		processors.add(new FreeMarkerRunner("word/document.xml"));

		docx.apply(processors, rootMap);

		File saveFile = new File(".test/chartTest-save.docx");
		docx.save(new FileOutputStream(saveFile));
		// tearDownHelper.add(saveFile);
	}

	@Test
	public void mainTest() throws IOException, JSONException {
		File targetDir = new File(".test/mainTest");
		targetDir.mkdirs();
		// tearDownHelper.add(targetDir);

		OOXMLPackage docx = new OOXMLPackage();
		docx.load(getClass().getResourceAsStream("/nestedList2.docx"), targetDir);

		InputStreamReader inputReader = new InputStreamReader(getClass().getResourceAsStream("/nestedListTest.in"));
		JSONTokener tokener = new JSONTokener(inputReader);
		Map<String, Object> rootMap = JsonHelper.parse((JSONObject) tokener.nextValue());

		List<OOXMLProcessor> processors = new ArrayList<OOXMLProcessor>();
		processors.add(new MergeFieldParser());
		processors.add(new AugmentedDirectiveProcessor());
		processors.add(new MagicNodeUnwrapper("word/document.xml"));
		processors.add(new FreeMarkerRunner("word/document.xml"));

		docx.apply(processors, rootMap);

		File saveFile = new File(".test/mainTest-save.docx");
		docx.save(new FileOutputStream(saveFile));
		// tearDownHelper.add(saveFile);
	}

	@Test
	public void saveTest() throws IOException {
		File targetDir = new File(".test/saveTest");
		targetDir.mkdirs();
		tearDownHelper.add(targetDir);

		OOXMLPackage docx = new OOXMLPackage();
		docx.load(getClass().getResourceAsStream("/nestedList2.docx"), targetDir);

		File saveFile = new File(".test/saveTest.docx");
		docx.save(new FileOutputStream(saveFile));
		tearDownHelper.add(saveFile);

		// diff word file
	}

	@Test
	public void extractingTest() throws IOException {
		File targetDir = new File(".test/extractingTest");
		targetDir.mkdirs();
		tearDownHelper.add(targetDir);

		OOXMLPackage docx = new OOXMLPackage();
		docx.load(getClass().getResourceAsStream("/extractingTest.docx"), targetDir);

		String[] relPaths = { "", "[Content_Types].xml", "_rels", "_rels\\.rels", "customXml", "customXml\\_rels",
				"customXml\\_rels\\item1.xml.rels", "customXml\\item1.xml", "customXml\\itemProps1.xml", "docProps",
				"docProps\\app.xml", "docProps\\core.xml", "word", "word\\_rels", "word\\_rels\\document.xml.rels",
				"word\\charts", "word\\charts\\_rels", "word\\charts\\_rels\\chart1.xml.rels",
				"word\\charts\\chart1.xml", "word\\document.xml", "word\\embeddings",
				"word\\embeddings\\Microsoft_Excel_____1.xlsx", "word\\endnotes.xml", "word\\fontTable.xml",
				"word\\footnotes.xml", "word\\settings.xml", "word\\styles.xml", "word\\stylesWithEffects.xml",
				"word\\theme", "word\\theme\\theme1.xml", "word\\webSettings.xml" };

		ArrayList<File> extractedFiles = new ArrayList<File>();
		ZipHelper.getFilesRecursivelyIn(targetDir, extractedFiles);

		File[] files = new File[relPaths.length];
		for (int i = 0; i < relPaths.length; ++i) {
			files[i] = new File(targetDir, relPaths[i]);
		}
		Arrays.sort(files);
		Object[] extractedArray = extractedFiles.toArray();
		Arrays.sort(extractedArray);

		assertTrue(Arrays.equals(extractedArray, files));
	}

	@Test
	public void relsTest() {
	}
}
