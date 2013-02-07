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

import java.io.File;
import java.util.ArrayList;

public class TearDownHelper {
	public ArrayList<File> outputFiles = new ArrayList<File>();

	public void tearDown() {
		for (File f : outputFiles) {
			try {
				if (f.isDirectory())
					deleteDir(f);
				else
					f.delete();
			} catch (Exception e) {
				// ignore
			}
		}
	}

	public static void deleteDir(File file) {
		if (file.isDirectory()) {
			if (file.listFiles().length != 0) {
				File[] fileList = file.listFiles();
				for (int i = 0; i < fileList.length; i++) {
					deleteDir(fileList[i]);
					file.delete();
				}
			} else {
				file.delete();
			}
		} else {
			file.delete();
		}
	}

	public void add(File file) {
		outputFiles.add(file);
	}

}
