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
import java.util.List;

import org.w3c.dom.NamedNodeMap;

public class Relationship {
	public Relationship parent;
	public String id;
	public String target;
	public String type;
	public List<Relationship> children;

	public Relationship() {
		parent = null;
		children = new ArrayList<Relationship>();
	}

	public Relationship(Relationship p, NamedNodeMap m) {
		parent = p;
		id = m.getNamedItem("Id").getNodeValue();
		target = m.getNamedItem("Target").getNodeValue();
		type = m.getNamedItem("Type").getNodeValue();
		children = new ArrayList<Relationship>();
	}
	
	public String toString() {
		if (type != null) {
			String prefix = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
			String shortType = type.startsWith(prefix) ? type.substring(prefix.length()) : type;
			if (children.isEmpty())
				return String.format("[%s, %s, %s]", id, target, shortType);
			else 
				return String.format("[%s, %s, %s, %d children]", id, target, shortType, children.size());
		} else {
			return String.format("[root, %d children]", children.size());
		}
		
	}
	
	public String toSummaryString() {
		if (target == null)
			return String.format("%s", toSummaryString(children));
		else {
			if (children == null || children.isEmpty()) {
				return String.format("%s", target);
			} else {
				return String.format("[%s%s]", target, toSummaryString(children));
			}
		}
	}

	private Object toSummaryString(List<Relationship> children) {
		if (children == null)
			return "";
		if (children.size() > 0) {
			StringBuilder sb = new StringBuilder();
			sb.append(" [");
			for (Relationship r: children) {
				if (r != children.get(0))
					sb.append(" ");
				sb.append(r.toSummaryString());
			}
			sb.append("]");
			return sb.toString();
		} else {
			return "";
		}
	}
}