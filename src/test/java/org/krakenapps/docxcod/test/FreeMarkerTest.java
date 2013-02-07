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

import static org.junit.Assert.assertEquals;

import java.io.InputStreamReader;
import java.io.StringWriter;
import java.util.List;
import java.util.Map;
import java.util.Scanner;

import org.json.JSONObject;
import org.json.JSONTokener;
import org.junit.Test;
import org.krakenapps.docxcod.JsonHelper;
import org.krakenapps.docxcod.util.CloseableHelper;

import freemarker.core.Environment;
import freemarker.template.Configuration;
import freemarker.template.DefaultObjectWrapper;
import freemarker.template.Template;
import freemarker.template.TemplateMethodModelEx;
import freemarker.template.TemplateModelException;

public class FreeMarkerTest {

	public class MakeNewChartFunction implements TemplateMethodModelEx {
		public int callCount = 0;
		
		@Override
		public Object exec(@SuppressWarnings("rawtypes") List arguments) throws TemplateModelException {
			callCount ++;
			System.out.printf("makeNewChart called(%s, %s)\n", arguments.get(0), arguments.get(1));
			System.out.printf("%s\n", Environment.getCurrentEnvironment().getKnownVariableNames());
			System.out.printf("%s\n", Environment.getCurrentEnvironment().getVariable("u"));
			return "";
		}
	}

	@Test
	public void UserDefMethodTest() throws Exception {
		InputStreamReader templateReader = null;
		InputStreamReader inputReader = null;
		Scanner scanner = null;
		try {
			Configuration cfg = new Configuration();
			cfg.setObjectWrapper(new DefaultObjectWrapper());

			inputReader = new InputStreamReader(getClass().getResourceAsStream("/nestedListTest.in"));
			JSONTokener tokener = new JSONTokener(inputReader);
			Map<String, Object> rootMap = JsonHelper.parse((JSONObject) tokener.nextValue());
			
			MakeNewChartFunction makeNewChartFunction = new MakeNewChartFunction();
			rootMap.put("makeNewChart", makeNewChartFunction);
			
			templateReader = new InputStreamReader(getClass().getResourceAsStream("/userDefMethodTest.fpl"));
			Template t = new Template("UserDefMethodTest", templateReader, cfg);

			StringWriter out = new StringWriter();

			t.process(rootMap, out);

			scanner = new Scanner(getClass().getResourceAsStream("/userDefMethodTest.out"));
			String expectedOutput = scanner.useDelimiter("\\A").next();
			
			assertEquals(expectedOutput, out.toString());
			assertEquals(3, makeNewChartFunction.callCount);
	
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			CloseableHelper.safeClose(inputReader);
			if (scanner != null)
				scanner.close();
		}
	}
}
