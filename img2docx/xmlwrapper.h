#pragma once
#import "MSXML3.dll" named_guids 
using namespace MSXML2;

#include <string>
using namespace std;

// --- CXmlNodeWrapper ---
class CXmlNodeWrapper  
{
private:
	bool m_bAutoRelease;
	MSXML2::IXMLDOMNodePtr m_xmlnode;
public:
	CXmlNodeWrapper();
	CXmlNodeWrapper(MSXML2::IXMLDOMNodePtr pNode,bool bAutoRelease = true);
	void operator=(MSXML2::IXMLDOMNodePtr pNode);
	MSXML2::IXMLDOMNode* Interface();
	bool IsValid();
	string Name();
	long ChildrenCount();
	MSXML2::IXMLDOMNodePtr GetChildNode(int nodeIndex);
	MSXML2::IXMLDOMNodePtr FindChildNode(string searchString);
	MSXML2::IXMLDOMNodePtr InsertChildAfter(MSXML2::IXMLDOMNodePtr refNode, string nodeName);
	MSXML2::IXMLDOMNodePtr InsertChildBefore(MSXML2::IXMLDOMNodePtr refNode, string nodeName);
	void SetAttrValue(string attrName, string value);
};

// --- CXmlDocWrapper ---
class CXmlDocWrapper
{
private:
	MSXML2::IXMLDOMDocumentPtr m_xmldoc;
public:
	CXmlDocWrapper();
	~CXmlDocWrapper();

	bool IsValid();
	bool Load(string file);
	bool Save(string file = "");
	MSXML2::IXMLDOMNodePtr AsNode();
};

