package com.silverio.android;

public interface IPersistable {

	public enum NormalizeMode {
		None,
		UTF8,
		UTF16
	}
	
	public void loadFromFile(String fileName);
	public void loadFromString(String s);
	public void saveToFile(String fileName, NormalizeMode mode);
	public void saveToString(String s, NormalizeMode mode);

}