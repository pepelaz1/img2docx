// XmlNodeWrapper.cpp: implementation of the CXmlNodeWrapper class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "XmlNodeWrapper.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CXmlNodeWrapper::CXmlNodeWrapper()
{
}

CXmlNodeWrapper::CXmlNodeWrapper(MSXML2::IXMLDOMNodePtr pNode,BOOL bAutoRelease)
{
	m_xmlnode = pNode;
	m_bAutoRelease = bAutoRelease;
}
void CXmlNodeWrapper::operator=(MSXML2::IXMLDOMNodePtr pNode)
{
	m_xmlnode = pNode;
}

CXmlNodeWrapper::~CXmlNodeWrapper()
{
	if (!m_bAutoRelease)
		m_xmlnode.Detach();
}


CString CXmlNodeWrapper::GetValue(LPCTSTR valueName)
{
	if (!IsValid())
		return "";

	MSXML2::IXMLDOMNodePtr attribute = m_xmlnode->Getattributes()->getNamedItem(valueName);
	if (attribute)
	{
		return (LPCSTR)attribute->Gettext();
	}
	return "";
}

BOOL CXmlNodeWrapper::IsValid()
{
	if (m_xmlnode == NULL)
		return FALSE;
	if (m_xmlnode.GetInterfacePtr() == NULL)
		return FALSE;
	return TRUE;
}

MSXML2::IXMLDOMNodePtr CXmlNodeWrapper::GetPrevSibling()
{
	if (!IsValid())
		return NULL;
	return m_xmlnode->GetpreviousSibling();
}

MSXML2::IXMLDOMNodePtr CXmlNodeWrapper::GetNextSibling()
{
	if (!IsValid())
		return NULL;
	return m_xmlnode->GetnextSibling();
}

MSXML2::IXMLDOMNodePtr CXmlNodeWrapper::GetNode(LPCTSTR nodeName)
{
	if (!IsValid())
		return NULL;
	try{
		return m_xmlnode->selectSingleNode(nodeName);
	}
	catch (_com_error e)
	{
		CString err = e.ErrorMessage();
	}
	return NULL;
}

MSXML2::IXMLDOMNodePtr CXmlNodeWrapper::GetNode(int nodeIndex)
{
	if (!IsValid())
		return NULL;
	return m_xmlnode->GetchildNodes()->Getitem(nodeIndex);
}

MSXML2::IXMLDOMNodePtr CXmlNodeWrapper::FindNode(LPCTSTR searchString)
{
	if (!IsValid())
		return NULL;
	try{
		return m_xmlnode->selectSingleNode(searchString);
	}
	catch (_com_error e)
	{
		CString err = e.ErrorMessage();
	}
	return NULL;
}

MSXML2::IXMLDOMNode* CXmlNodeWrapper::Detach()
{
	if (IsValid())
	{
		return m_xmlnode.Detach();
	}
	else
		return NULL;
}

long CXmlNodeWrapper::NumNodes()
{
	if (IsValid())
	{
		return m_xmlnode->GetchildNodes()->Getlength();
	}
	else
		return 0;
}

void CXmlNodeWrapper::SetValue(LPCTSTR valueName,LPCTSTR value)
{
	MSXML2::IXMLDOMDocumentPtr xmlDocument = m_xmlnode->GetownerDocument();
	if (xmlDocument)
	{
		MSXML2::IXMLDOMNamedNodeMapPtr attributes = m_xmlnode->Getattributes();
		if (attributes)
		{
			MSXML2::IXMLDOMAttributePtr attribute = xmlDocument->createAttribute(valueName);
			if (attribute)
			{
				attribute->Puttext(value);
				attributes->setNamedItem(attribute);
			}
		}
	}
}

void CXmlNodeWrapper::SetValue(LPCTSTR valueName,int value)
{
	CString str;
	str.Format("%ld",value);
	SetValue(valueName,str);
}

void CXmlNodeWrapper::SetValue(LPCTSTR valueName,short value)
{
	CString str;
	str.Format("%hd",value);
	SetValue(valueName,str);
}

void CXmlNodeWrapper::SetValue(LPCTSTR valueName,double value)
{
	CString str;
	str.Format("%f",value);
	SetValue(valueName,str);
}

void CXmlNodeWrapper::SetValue(LPCTSTR valueName,float value)
{
	CString str;
	str.Format("%f",value);
	SetValue(valueName,str);
}

void CXmlNodeWrapper::SetValue(LPCTSTR valueName,bool value)
{
	CString str;
	if (value)
		str = "True";
	else
		str = "False";
	SetValue(valueName,str);
}

MSXML2::IXMLDOMNodePtr CXmlNodeWrapper::InsertNode(int index,LPCTSTR nodeName)
{
	MSXML2::IXMLDOMDocumentPtr xmlDocument = m_xmlnode->GetownerDocument();
	if (xmlDocument)
	{
		MSXML2::IXMLDOMNodePtr newNode = xmlDocument->createNode(_variant_t((short)MSXML2::NODE_ELEMENT),nodeName,"");
		MSXML2::IXMLDOMNodePtr refNode = GetNode(index);
		if (refNode)
			newNode = m_xmlnode->insertBefore(newNode,_variant_t(refNode.GetInterfacePtr()));
		else
			newNode = m_xmlnode->appendChild(newNode);
		return newNode;
	}
	return NULL;
}

MSXML2::IXMLDOMNodePtr CXmlNodeWrapper::InsertNode(int index,MSXML2::IXMLDOMNodePtr pNode)
{
	MSXML2::IXMLDOMNodePtr newNode = pNode->cloneNode(VARIANT_TRUE);
	if (newNode)
	{
		MSXML2::IXMLDOMNodePtr refNode = GetNode(index);
		if (refNode)
			newNode = m_xmlnode->insertBefore(newNode,_variant_t(refNode.GetInterfacePtr()));
		else
			newNode = m_xmlnode->appendChild(newNode);
		return newNode;
	}
	else
		return NULL;
}

CString CXmlNodeWrapper::GetXML()
{
	if (IsValid())
		return (LPCSTR)m_xmlnode->Getxml();
	else
		return "";
}

MSXML2::IXMLDOMNodePtr CXmlNodeWrapper::RemoveNode(MSXML2::IXMLDOMNodePtr pNode)
{
	if (!IsValid())
		return NULL;
	return m_xmlnode->removeChild(pNode);
}

/* ********************************************************************************************************* */
CXmlDocumentWrapper::CXmlDocumentWrapper()
{
	m_xmldoc.CreateInstance(MSXML2::CLSID_DOMDocument);
}

CXmlDocumentWrapper::CXmlDocumentWrapper(MSXML2::IXMLDOMDocumentPtr pDoc)
{
	m_xmldoc = pDoc;
}

void CXmlDocumentWrapper::operator=(MSXML2::IXMLDOMDocumentPtr pDoc)
{
	if (IsValid())
		m_xmldoc.Release();
	m_xmldoc = pDoc;
}


CXmlDocumentWrapper::~CXmlDocumentWrapper()
{
}

BOOL CXmlDocumentWrapper::IsValid()
{
	if (m_xmldoc == NULL)
		return FALSE;
	if (m_xmldoc.GetInterfacePtr() == NULL)
		return FALSE;
	return TRUE;
}

MSXML2::IXMLDOMDocument* CXmlDocumentWrapper::Detach()
{
	if (!IsValid())
		return NULL;
	return m_xmldoc.Detach();
}

MSXML2::IXMLDOMDocumentPtr CXmlDocumentWrapper::Clone()
{
	if (!IsValid())
		return NULL;
	MSXML2::IXMLDOMDocumentPtr xmldoc;
	xmldoc.CreateInstance(MSXML2::CLSID_DOMDocument);
	_variant_t v(xmldoc.GetInterfacePtr());
	m_xmldoc->save(v);
	return xmldoc;
}

MSXML2::IXMLDOMDocument* CXmlDocumentWrapper::Interface()
{
	if (IsValid())
		return m_xmldoc;
	return NULL;
}

BOOL CXmlDocumentWrapper::Load(LPCTSTR path)
{
	if (!IsValid())
		return FALSE;

	_variant_t v(path);
	m_xmldoc->put_async(VARIANT_FALSE);
	VARIANT_BOOL success = m_xmldoc->load(v);
	if (success == VARIANT_TRUE)
		return TRUE;
	else
		return FALSE;
}

BOOL CXmlDocumentWrapper::LoadXML(LPCTSTR xml)
{
	if (!IsValid())
		return FALSE;
	VARIANT_BOOL success = m_xmldoc->loadXML(xml);
	if (success == VARIANT_TRUE)
		return TRUE;
	else
		return FALSE;
}

BOOL CXmlDocumentWrapper::Save(LPCTSTR path)
{
	try
	{
		if (!IsValid())
			return FALSE;
		CString szPath(path);
		if (szPath == "")
		{
			_bstr_t curPath = m_xmldoc->Geturl();
			szPath = (LPSTR)curPath;
		}
		_variant_t v(szPath);
		if (FAILED(m_xmldoc->save(v)))
			return FALSE;
		else
			return TRUE;
	}
	catch(...)
	{
		return FALSE;
	}
}

MSXML2::IXMLDOMNodePtr CXmlDocumentWrapper::AsNode()
{
	if (!IsValid())
		return NULL;
	return m_xmldoc->GetdocumentElement();
}

CString CXmlDocumentWrapper::GetXML()
{
	if (IsValid())
		return (LPCSTR)m_xmldoc->Getxml();
	else
		return "";
}

CString CXmlDocumentWrapper::GetUrl()
{
	return (LPSTR)m_xmldoc->Geturl();
}

MSXML2::IXMLDOMDocumentPtr CXmlNodeWrapper::ParentDocument()
{
	return m_xmlnode->GetownerDocument();
}

MSXML2::IXMLDOMNode* CXmlNodeWrapper::Interface()
{
	if (IsValid())
		return m_xmlnode;
	return NULL;
}

MSXML2::IXMLDOMNodePtr CXmlNodeWrapper::InsertBefore(MSXML2::IXMLDOMNodePtr refNode, LPCTSTR nodeName)
{
	MSXML2::IXMLDOMDocumentPtr xmlDocument = m_xmlnode->GetownerDocument();
	if (xmlDocument)
	{
		MSXML2::IXMLDOMNodePtr newNode = xmlDocument->createNode(_variant_t((short)MSXML2::NODE_ELEMENT),nodeName,"");
		newNode = m_xmlnode->insertBefore(newNode,_variant_t(refNode.GetInterfacePtr()));
		return newNode;
	}
	return NULL;
}

MSXML2::IXMLDOMNodePtr CXmlNodeWrapper::InsertAfter(MSXML2::IXMLDOMNodePtr refNode, LPCTSTR nodeName)
{
	MSXML2::IXMLDOMDocumentPtr xmlDocument = m_xmlnode->GetownerDocument();
	if (xmlDocument)
	{
		MSXML2::IXMLDOMNodePtr newNode = xmlDocument->createNode(_variant_t((short)MSXML2::NODE_ELEMENT),nodeName,"");
		MSXML2::IXMLDOMNodePtr nextNode = refNode->GetnextSibling();
		if (nextNode.GetInterfacePtr() != NULL)
			newNode = m_xmlnode->insertBefore(newNode,_variant_t(nextNode.GetInterfacePtr()));
		else
			newNode = m_xmlnode->appendChild(newNode);
		return newNode;
	}
	return NULL;
}

void CXmlNodeWrapper::RemoveNodes(LPCTSTR searchStr)
{
	if (!IsValid())
		return;
	MSXML2::IXMLDOMNodeListPtr nodeList = m_xmlnode->selectNodes(searchStr);
	for (int i = 0; i < nodeList->Getlength(); i++)
	{
		try
		{
			MSXML2::IXMLDOMNodePtr pNode = nodeList->Getitem(i);
			pNode->GetparentNode()->removeChild(pNode);
		}
		catch (_com_error er)
		{
			AfxMessageBox(er.ErrorMessage());
		}
	}
}

MSXML2::IXMLDOMNodePtr CXmlNodeWrapper::Parent()
{
	if (IsValid())
		return m_xmlnode->GetparentNode();
	return NULL;
}

CXmlNodeListWrapper::CXmlNodeListWrapper()
{
}
CXmlNodeListWrapper::CXmlNodeListWrapper(MSXML2::IXMLDOMNodeListPtr pList)
{
	m_xmlnodelist = pList;
}

void CXmlNodeListWrapper::operator=(MSXML2::IXMLDOMNodeListPtr pList)
{
	m_xmlnodelist = pList;
}

CXmlNodeListWrapper::~CXmlNodeListWrapper()
{
	
}

int CXmlNodeListWrapper::Count()
{
	if (IsValid())
		return m_xmlnodelist->Getlength();
	else
		return 0;
	
}

BOOL CXmlNodeListWrapper::IsValid()
{
	if (m_xmlnodelist == NULL)
		return FALSE;
	if (m_xmlnodelist.GetInterfacePtr() == NULL)
		return FALSE;
	return TRUE;
}

MSXML2::IXMLDOMNodePtr CXmlNodeListWrapper::Next()
{
	if (IsValid())
		return m_xmlnodelist->nextNode();
	else
		return NULL;
}

void CXmlNodeListWrapper::Start()
{
	if (IsValid())
		m_xmlnodelist->reset();
}

MSXML2::IXMLDOMNodePtr CXmlNodeListWrapper::Node(int index)
{
	if (IsValid())
		return m_xmlnodelist->Getitem(index);
	else
		return NULL;
}

MSXML2::IXMLDOMDocumentPtr CXmlNodeListWrapper::AsDocument()
{
	if (IsValid())
	{
		CXmlDocumentWrapper doc;
		doc.LoadXML("<NodeList></NodeList>");
		CXmlNodeWrapper root(doc.AsNode());
		
		for (int i = 0; i < m_xmlnodelist->Getlength(); i++)
		{
			root.InsertNode(root.NumNodes(),m_xmlnodelist->Getitem(i)->cloneNode(VARIANT_TRUE));
		}
		return doc.Interface();
	}
	else 
		return NULL;
}

MSXML2::IXMLDOMNodeListPtr CXmlNodeWrapper::FindNodes(LPCTSTR searchStr)
{
	if(IsValid())
	{
		try{
			return m_xmlnode->selectNodes(searchStr);
		}
		catch (_com_error e)
		{
			CString err = e.ErrorMessage();
			return NULL;
		}
	}
	else
		return NULL;
}

CString CXmlNodeWrapper::Name()
{
	if (IsValid())
		return (LPCSTR)m_xmlnode->GetbaseName();
	return "";
}

MSXML2::IXMLDOMNodePtr CXmlNodeWrapper::InsertAfter(MSXML2::IXMLDOMNodePtr refNode, MSXML2::IXMLDOMNodePtr pNode)
{
	MSXML2::IXMLDOMNodePtr nextNode = refNode->GetnextSibling();
	MSXML2::IXMLDOMNodePtr newNode;
	if (nextNode.GetInterfacePtr() != NULL)
		newNode = m_xmlnode->insertBefore(pNode,_variant_t(nextNode.GetInterfacePtr()));
	else
		newNode = m_xmlnode->appendChild(pNode);
	return newNode;
}

void CXmlNodeWrapper::SetText(LPCTSTR text)
{
	if (IsValid())
		m_xmlnode->Puttext(text);
}

CString CXmlNodeWrapper::GetText()
{
	if (IsValid())
		return (LPCSTR)m_xmlnode->Gettext();
	else
		return "";
}

void CXmlNodeWrapper::ReplaceNode(MSXML2::IXMLDOMNodePtr pOldNode, MSXML2::IXMLDOMNodePtr pNewNode)
{
	if (IsValid())
	{
		m_xmlnode->replaceChild(pNewNode,pOldNode);
	}
}

int CXmlNodeWrapper::NumAttributes()
{
	if (IsValid())
	{
		MSXML2::IXMLDOMNamedNodeMapPtr attribs = m_xmlnode->Getattributes();
		if (attribs)
			return attribs->Getlength();
	}

	return 0;
}

CString CXmlNodeWrapper::GetAttribName(int index)
{
	if (IsValid())
	{
		MSXML2::IXMLDOMNamedNodeMapPtr attribs = m_xmlnode->Getattributes();
		if (attribs)
		{
			MSXML2::IXMLDOMAttributePtr attrib = attribs->Getitem(index);
			if (attrib)
				return (LPCSTR)attrib->Getname();
		}
	}

	return "";
}

CString CXmlNodeWrapper::GetAttribVal(int index)
{
	if (IsValid())
	{
		MSXML2::IXMLDOMNamedNodeMapPtr attribs = m_xmlnode->Getattributes();
		if (attribs)
		{
			MSXML2::IXMLDOMAttributePtr attrib = attribs->Getitem(index);
			if (attrib)
				return (LPCSTR)attrib->Gettext();
		}
	}

	
	return "";
}

CString CXmlNodeWrapper::NodeType()
{
	if (IsValid())
		return (LPCSTR)m_xmlnode->GetnodeTypeString();
	return "";
}
