// XmlNodeWrapper.h: interface for the CXmlNodeWrapper class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_XMLNODEWRAPPER_H__43622334_FDEB_4175_9E6D_19BBAA3992A5__INCLUDED_)
#define AFX_XMLNODEWRAPPER_H__43622334_FDEB_4175_9E6D_19BBAA3992A5__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#import "MSXML3.dll" named_guids 
using namespace MSXML2;

class CXmlNodeWrapper  
{
public:
	CString NodeType();
	CString GetAttribVal(int index);
	CString GetAttribName(int index);
	int NumAttributes();
	void ReplaceNode(MSXML2::IXMLDOMNodePtr pOldNode,MSXML2::IXMLDOMNodePtr pNewNode);
	CString GetText();
	void SetText(LPCTSTR text);
	MSXML2::IXMLDOMNodePtr InsertAfter(MSXML2::IXMLDOMNodePtr refNode, MSXML2::IXMLDOMNodePtr pNode);
	CString Name();
	MSXML2::IXMLDOMNodeListPtr FindNodes(LPCTSTR searchStr);
	MSXML2::IXMLDOMNodePtr Parent();
	void RemoveNodes(LPCTSTR searchStr);
	MSXML2::IXMLDOMNodePtr InsertAfter(MSXML2::IXMLDOMNodePtr refNode, LPCTSTR nodeName);
	MSXML2::IXMLDOMNodePtr InsertBefore(MSXML2::IXMLDOMNodePtr refNode, LPCTSTR nodeName);
	MSXML2::IXMLDOMNode* Interface();
	MSXML2::IXMLDOMDocumentPtr ParentDocument();
	CString GetXML();
	MSXML2::IXMLDOMNodePtr RemoveNode(MSXML2::IXMLDOMNodePtr pNode);
	MSXML2::IXMLDOMNodePtr InsertNode(int index,LPCTSTR nodeName);
	MSXML2::IXMLDOMNodePtr InsertNode(int index,MSXML2::IXMLDOMNodePtr pNode);
	long NumNodes();
	MSXML2::IXMLDOMNode* Detach();
	MSXML2::IXMLDOMNodePtr GetNode(LPCTSTR nodeName);
	MSXML2::IXMLDOMNodePtr GetNode(int nodeIndex);
	MSXML2::IXMLDOMNodePtr FindNode(LPCTSTR searchString);
	MSXML2::IXMLDOMNodePtr GetPrevSibling();
	MSXML2::IXMLDOMNodePtr GetNextSibling();
	void SetValue(LPCTSTR valueName,LPCTSTR value);
	void SetValue(LPCTSTR valueName,int value);
	void SetValue(LPCTSTR valueName,short value);
	void SetValue(LPCTSTR valueName,double value);
	void SetValue(LPCTSTR valueName,float value);
	void SetValue(LPCTSTR valueName,bool value);
	BOOL IsValid();
	CString GetValue(LPCTSTR valueName);

	CXmlNodeWrapper();
	CXmlNodeWrapper(MSXML2::IXMLDOMNodePtr pNode,BOOL bAutoRelease = TRUE);
	void operator=(MSXML2::IXMLDOMNodePtr pNode);
	virtual ~CXmlNodeWrapper();
private:
	BOOL m_bAutoRelease;
	MSXML2::IXMLDOMNodePtr m_xmlnode;
};

class CXmlDocumentWrapper
{
public:
	CString GetUrl();
	CString GetXML();
	BOOL IsValid();
	BOOL Load(LPCTSTR path);
	BOOL LoadXML(LPCTSTR xml);
	BOOL Save(LPCTSTR path = "");
	MSXML2::IXMLDOMDocument* Detach();
	MSXML2::IXMLDOMDocumentPtr Clone();
	MSXML2::IXMLDOMDocument* Interface();
	CXmlDocumentWrapper();
	CXmlDocumentWrapper(MSXML2::IXMLDOMDocumentPtr pDoc);
	void operator=(MSXML2::IXMLDOMDocumentPtr pDoc);
	MSXML2::IXMLDOMNodePtr AsNode();
	virtual ~CXmlDocumentWrapper();
private:
	MSXML2::IXMLDOMDocumentPtr m_xmldoc;
};

class CXmlNodeListWrapper
{
public:
	MSXML2::IXMLDOMDocumentPtr AsDocument();
	MSXML2::IXMLDOMNodePtr Node(int index);
	void Start();
	MSXML2::IXMLDOMNodePtr Next();
	BOOL IsValid();
	int Count();
	CXmlNodeListWrapper();
	CXmlNodeListWrapper(MSXML2::IXMLDOMNodeListPtr pList);
	void operator=(MSXML2::IXMLDOMNodeListPtr pList);
	virtual ~CXmlNodeListWrapper();

private:
	MSXML2::IXMLDOMNodeListPtr m_xmlnodelist;
};

#endif // !defined(AFX_XMLNODEWRAPPER_H__43622334_FDEB_4175_9E6D_19BBAA3992A5__INCLUDED_)
