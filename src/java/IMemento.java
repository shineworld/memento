package com.silverio.android;

import java.util.UUID;

public interface IMemento {

	public IMemento createChild(String type);
	public IMemento createChildSmart(String type);
	public Boolean getBoolean(String key);
	public IMemento getChild(String type);
	public IMemento[] getChildren(String type);
	public IMemento getChildFromPath(String path);
	public IMemento[] getChildrenFromPath(String path);
	public Double getDouble(String key);
	public UUID getGUID(String key);
	public Integer getInteger(String key);
	public String getName();
	public IMemento getRoot();
	public IMemento getParent();
	public String getString(String key);
	public String getTextData(String value);
	public void putBoolean(String key, boolean value);
	public void putDouble(String key, double value);
	public void putGUID(String key, UUID value);
	public void putInteger(String key, int value);
	public void putString(String key, String value);
	public void putTextData(String value);

}
