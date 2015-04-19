#include "xmlwrapper.h"


// --- CXmlNodeWrapper ---
CXmlNodeWrapper::CXmlNodeWrapper()
{
}

CXmlNodeWrapper::CXmlNodeWrapper(MSXML2::IXMLDOMNodePtr pNode,bool bAutoRelease)
{
	m_xmlnode = pNode;
	m_bAutoRelease = bAutoRelease;
}

void CXmlNodeWrapper::operator=(MSXML2::IXMLDOMNodePtr pNode)
{
	m_xmlnode = pNode;
}

bool CXmlNodeWrapper::IsValid()
{
	if (m_xmlnode == NULL)
		return false;
	if (m_xmlnode.GetInterfacePtr() == NULL)
		return FALSE;
	return true;
}

MSXML2::IXMLDOMNode* CXmlNodeWrapper::Interface()
{
	if (IsValid())
		return m_xmlnode;
	return NULL;

}
string CXmlNodeWrapper::Name()
{
	if (IsValid())
		return (LPCSTR)m_xmlnode->GetbaseName();
	return "";
}

long CXmlNodeWrapper::ChildrenCount()
{
	if (IsValid())
	{
		return m_xmlnode->GetchildNodes()->Getlength();
	}
	else
		return 0;
}

MSXML2::IXMLDOMNodePtr CXmlNodeWrapper::GetChildNode(int nodeIndex)
{
	if (!IsValid())
		return NULL;
	return m_xmlnode->GetchildNodes()->Getitem(nodeIndex);
}

MSXML2::IXMLDOMNodePtr CXmlNodeWrapper::FindChildNode(string searchString)
{
	if (!IsValid())
		return NULL;
	try{
		return m_xmlnode->selectSingleNode(searchString.data());
	}
	catch (_com_error e)
	{
		string err = e.ErrorMessage();
	}
	return NULL;
}

MSXML2::IXMLDOMNodePtr CXmlNodeWrapper::InsertChildBefore(MSXML2::IXMLDOMNodePtr refNode, string nodeName)
{
	MSXML2::IXMLDOMDocumentPtr xmlDocument = m_xmlnode->GetownerDocument();
	if (xmlDocument)
	{
		MSXML2::IXMLDOMNodePtr newNode = xmlDocument->createNode(_variant_t((short)MSXML2::NODE_ELEMENT),nodeName.c_str(),refNode->GetparentNode()->GetnamespaceURI());
		newNode = m_xmlnode->insertBefore(newNode,_variant_t(refNode.GetInterfacePtr()));

		return newNode;
	}
	return NULL;
}

MSXML2::IXMLDOMNodePtr CXmlNodeWrapper::InsertChildAfter(MSXML2::IXMLDOMNodePtr refNode, string nodeName)
{
	MSXML2::IXMLDOMDocumentPtr xmlDocument = m_xmlnode->GetownerDocument();
	if (xmlDocument)
	{
		MSXML2::IXMLDOMNodePtr newNode = xmlDocument->createNode(_variant_t((short)MSXML2::NODE_ELEMENT),nodeName.c_str(),"");
		MSXML2::IXMLDOMNodePtr nextNode = refNode->GetnextSibling();
		if (nextNode.GetInterfacePtr() != NULL)
			newNode = m_xmlnode->insertBefore(newNode,_variant_t(nextNode.GetInterfacePtr()));
		else
			newNode = m_xmlnode->appendChild(newNode);
		return newNode;
	}
	return NULL;
}

void CXmlNodeWrapper::SetAttrValue(string attrName,string value)
{
	MSXML2::IXMLDOMDocumentPtr xmlDocument = m_xmlnode->GetownerDocument();
	if (xmlDocument)
	{
		MSXML2::IXMLDOMNamedNodeMapPtr attributes = m_xmlnode->Getattributes();
		if (attributes)
		{
			MSXML2::IXMLDOMAttributePtr attribute = xmlDocument->createAttribute(attrName.c_str());
			if (attribute)
			{
				attribute->Puttext(value.c_str());
				attributes->setNamedItem(attribute);
			}
		}
	}
}

// --- CXmlDocWrapper ---
CXmlDocWrapper::CXmlDocWrapper()
{
	CoInitialize(NULL);
	m_xmldoc.CreateInstance(MSXML2::CLSID_DOMDocument);
}


CXmlDocWrapper::~CXmlDocWrapper()
{
	//CoUninitialize();
}

bool CXmlDocWrapper::IsValid()
{
	if (m_xmldoc == NULL)
		return false;
	if (m_xmldoc.GetInterfacePtr() == NULL)
		return false;
	return true;
}

bool CXmlDocWrapper::Load(string file)
{
	if (!IsValid())
		return false;

	_variant_t v(file.c_str());
	m_xmldoc->put_async(VARIANT_FALSE);
	VARIANT_BOOL success = m_xmldoc->load(v);
	if (success == VARIANT_TRUE)
		return true;
	else
		return false;
}

bool CXmlDocWrapper::Save(string file)
{
	try
	{
		if (!IsValid())
			return false;

		string sfile = file;
		if (sfile.empty())
			sfile = m_xmldoc->Geturl();
		_variant_t v(sfile.c_str());
		if (FAILED(m_xmldoc->save(v)))
			return false;
		else
			return true;
	}
	catch(...)
	{
		return false;
	}
}

MSXML2::IXMLDOMNodePtr CXmlDocWrapper::AsNode()
{
	if (!IsValid())
		return NULL;
	return m_xmldoc->GetdocumentElement();
}


