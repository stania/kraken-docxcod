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

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;

public class JsonHelper {
	public static Map<String, Object> parse(JSONObject obj) {
		Map<String, Object> m = new HashMap<String, Object>();
		String[] names = JSONObject.getNames(obj);
		if (names == null)
			return m;

		for (String key : names) {
			try {
				Object value = obj.get(key);
				if (value == JSONObject.NULL)
					value = null;
				else if (value instanceof JSONArray)
					value = parse((JSONArray) value);
				else if (value instanceof JSONObject)
					value = parse((JSONObject) value);

				m.put(key, value);
			} catch (JSONException e) {
				e.printStackTrace();
			}
		}

		return m;
	}

	public static List<Object> parse(JSONArray arr) {
		List<Object> list = new ArrayList<Object>();
		for (int i = 0; i < arr.length(); i++) {
			try {
				Object o = arr.get(i);
				if (o == JSONObject.NULL)
					list.add(null);
				else if (o instanceof JSONArray)
					list.add(parse((JSONArray) o));
				else if (o instanceof JSONObject)
					list.add(parse((JSONObject) o));
				else
					list.add(o);
			} catch (JSONException e) {
				e.printStackTrace();
			}
		}
		return list;
	}


}
