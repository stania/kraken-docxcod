package org.krakenapps.docxcod;

public class CellData implements Comparable<CellData> {
	public String address;
	public String data;

	public CellData(String address, String data) {
		this.address = address;
		this.data = data;
	}

	@Override
	public int compareTo(CellData o) {
		return address.compareTo(o.address);
	}
}
