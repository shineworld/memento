package com.silverio.android;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.util.ArrayList;
import java.util.UUID;

import org.w3c.dom.*;

import javax.xml.parsers.*;

public final class XMLMemento implements IMemento, IPersistable {
	private Document mDocument;
	private Element mElement;

	public static IMemento createReadRoot(String fileName) {
		try {
			DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
			DocumentBuilder parser = factory.newDocumentBuilder();
			Document document = parser.parse(new File(fileName));
			Element element = document.getDocumentElement();
			return new XMLMemento(document, element);
		} catch (Exception e) {
			return null;
		}
	}

	public static IMemento createWriteRoot(String type) {
		try {
			Document document = DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument();
			Element element = document.createElement(type);
			document.appendChild(element);
			return new XMLMemento(document, element);
		} catch (Exception e) {
			return null;
		}
	}

	public XMLMemento(Document document, Element element) {
		super();
		mDocument = document;
		mElement = element;
	}

	/* IMemento */

	@Override
	public IMemento createChild(String type) {
		Element child = mDocument.createElement(type);
		mElement.appendChild(child);
		return new XMLMemento(mDocument, child);
	}

	// public IMemento createChild(String type, String id)

	@Override
	public IMemento createChildSmart(String type) {
		IMemento result = getChild(type);
		if (result == null)
			result = createChild(type);
		return result;
	}

	// public IMemento copyChild(IMemento child)

	@Override
	public Boolean getBoolean(String key) {
		String value = getString(key);
		if (value == null)
			return null;
		try {
			return new Boolean(value);
		} catch (Exception e) {
			return null;
		}
	}

	@Override
	public IMemento getChild(String type) {

		// get nodes
		NodeList nodes = mElement.getChildNodes();
		int size = nodes.getLength();
		if (size == 0) {
			return null;
		}

		// find first node which is a child of this node
		for (int i = 0; i < size; i++) {
			Node node = nodes.item(i);
			if (node instanceof Element) {
				Element element = (Element) node;
				if (element.getNodeName().equals(type)) {
					return new XMLMemento(mDocument, element);
				}
			}
		}

		// a child was not found
		return null;
	}

	@Override
	public IMemento[] getChildren(String type) {

		// get nodes
		NodeList nodes = mElement.getChildNodes();
		int size = nodes.getLength();
		if (size == 0) {
			return new IMemento[0];
		}

		// extract each node with given type
		ArrayList list = new ArrayList(size);
		for (int i = 0; i < size; i++) {
			Node node = nodes.item(i);
			if (node instanceof Element) {
				Element element = (Element) node;
				if (element.getNodeName().equals(type)) {
					list.add(element);
				}
			}
		}

		// create a memento for each node
		size = list.size();
		IMemento[] results = new IMemento[size];
		for (int i = 0; i < size; i++) {
			results[i] = new XMLMemento(mDocument, (Element) list.get(i));
		}
		return results;
	}

	@Override
	public IMemento getChildFromPath(String path) {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public IMemento[] getChildrenFromPath(String path) {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public Double getDouble(String key) {
		String value = getString(key);
		if (value == null)
			return null;
		try {
			return new Double(value);
		} catch (Exception e) {
			return null;
		}
	}

	@Override
	public UUID getGUID(String key) {
		String value = getString(key);
		if (value == null)
			return null;
		try {
			if (value.startsWith("{") && value.endsWith("}"))
				value = new String(value.substring(1, value.length() - 1));
			UUID result = UUID.fromString(value.toLowerCase());
			return result;
		} catch (Exception e) {
			return null;
		}
	}

	@Override
	public Integer getInteger(String key) {
		String value = getString(key);
		if (value == null)
			return null;
		try {
			return new Integer(value);
		} catch (Exception e) {
			return null;
		}
	}

	@Override
	public String getName() {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public IMemento getRoot() {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public IMemento getParent() {
		try {
			return new XMLMemento(mDocument, (Element) mElement.getParentNode());
		} catch (Exception e) {
			return null;
		}
	}

	@Override
	public String getString(String key) {
		Attr attr = mElement.getAttributeNode(key);
		if (attr == null) {
			return null;
		}
		return attr.getValue();
	}

	@Override
	public String getTextData(String value) {
		// TODO Auto-generated method stub
		return null;
	}

	@Override
	public void putBoolean(String key, boolean value) {
		// TODO Auto-generated method stub

	}

	@Override
	public void putDouble(String key, double value) {
		// TODO Auto-generated method stub

	}

	@Override
	public void putGUID(String key, UUID value) {
		// TODO Auto-generated method stub

	}

	@Override
	public void putInteger(String key, int value) {
		// TODO Auto-generated method stub

	}

	@Override
	public void putString(String key, String value) {
		// TODO Auto-generated method stub

	}

	@Override
	public void putTextData(String value) {
		// TODO Auto-generated method stub

	}

	/* IPersistable */

	@Override
	public void loadFromFile(String FileName) {
		// TODO Auto-generated method stub

	}

	@Override
	public void loadFromString(String S) {
		try {
			DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
			DocumentBuilder parser = factory.newDocumentBuilder();
			mDocument = parser.parse(new ByteArrayInputStream(S.getBytes()));
			mElement = mDocument.getDocumentElement();
		} catch (Exception e) {
			// todo: ???
		}
	}

	@Override
	public void saveToFile(String FileName, NormalizeMode Mode) {
		// TODO Auto-generated method stub

	}

	@Override
	public void saveToString(String S, NormalizeMode Mode) {
		// TODO Auto-generated method stub

	}

}