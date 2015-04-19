#include "docxgen.h"
#include "resource.h"
#include "template.h"
#include <windows.h>
#include <fstream>
#include <sstream>
#include "filehelper.h"
#include "ziputils/zip.h"
#include "ziputils/unzip.h"



CDocxGenerator::CDocxGenerator(void)
{
	PopulateContentTypes();	
}


CDocxGenerator::~CDocxGenerator(void)
{
}

void CDocxGenerator::Reset(void)
{
	m_srcfiles.clear();
}

void CDocxGenerator::PopulateContentTypes()
{
	m_content_types["jpg"] = "image/jpeg";
}

void CDocxGenerator::Add(char *srcfile)
{
	m_srcfiles.push_back(srcfile);
}

void CDocxGenerator::ReadTemplate(void)
{
	int idx = 0;
	int x;   
	for ( unsigned int i = 0; i < 13; i++)
	{
		for ( unsigned int j = 0; j < strlen(g_template[i]); j+=2)
		{
			char sym[2];
			strncpy(sym,g_template[i]+j,2);
			sscanf_s(sym,"%X",&x);
			m_tbuff[idx++] = x;
		}
	}
}

void CDocxGenerator::UnzipTemplate(string path)
{
	char oldcurdir[MAX_PATH];
	GetCurrentDirectory(MAX_PATH, oldcurdir);
	SetCurrentDirectory(path.c_str());

	HZIP hz = OpenZip(m_tbuff,TEMPLATE_SIZE,0);
    ZIPENTRY ze;
	GetZipItem(hz,-1,&ze); 
	int numitems=ze.index;
	for (int zi=0; zi<numitems; zi++)
	{ 
		ZIPENTRY ze; 
		GetZipItem(hz, zi, &ze);
		UnzipItem(hz, zi, ze.name);       
	}
	CloseZip(hz);
	
	SetCurrentDirectory(oldcurdir);
}

void ZipRecursively(string root, string path, HZIP hz)
{
	string dir = path + "\\*";
	WIN32_FIND_DATA ffdata;
	HANDLE find = FindFirstFile(dir.c_str(), &ffdata);
	do
	{
		if (strcmp(ffdata.cFileName,".") !=0 && 
			strcmp(ffdata.cFileName,"..") != 0 &&
			strcmp(ffdata.cFileName,"template.docx") != 0)
		{
			if ( ffdata.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)
			{
				string cdir = path + "\\" + ffdata.cFileName;
				ZipRecursively(root, cdir,hz);
			}
			else
			{
				string dfile = path + "\\" + ffdata.cFileName;
				dfile.erase(0, root.length()+1);
				ZipAdd(hz,(const TCHAR *)dfile.c_str(),(const TCHAR *)dfile.c_str());
			}
		}
	}
	while(FindNextFileA(find, &ffdata));
	FindClose(find);
}

void CDocxGenerator::ZipTemplate(string path)
{
	char oldcurdir[MAX_PATH];
	GetCurrentDirectory(MAX_PATH, oldcurdir);
	SetCurrentDirectory(path.c_str());

	HZIP hz = CreateZip("template.docx",0);
	ZipRecursively(path, path, hz);
	CloseZip(hz);

	SetCurrentDirectory(oldcurdir);
}

void CDocxGenerator::CopyImages(string path)
{
	string imgpath = path + "\\word\\media";
	CreateDirectory(imgpath.c_str(), NULL);
	
	int n = 1;
	list<string>::const_iterator it;
    for(it = m_srcfiles.begin(); it != m_srcfiles.end(); it++)
	{
		stringstream stream;
		stream << "image";
		stream << n++;
		stream << ".";
		stream << GetFileExtension( *it);
		string imgfile = imgpath + "\\" +stream.str();
		CopyFile((*it).c_str(), imgfile.c_str(), false);
	}
	
}

void CDocxGenerator::MakeXMLChanges(string path)
{
	AddContentType(path);
}
void CDocxGenerator::AddContentType(string path)
{
	m_xml.Load(path+"\\[Content_Types].xml");
	CXmlNodeWrapper root(m_xml.AsNode());
	CXmlNodeWrapper bnode(root.FindChildNode("Override"));

	map<string,bool> unique_exts;
	list<string>::const_iterator it;
	for(it = m_srcfiles.begin(); it != m_srcfiles.end(); it++)
	{
		string ext = GetFileExtension(*it);
		if (!unique_exts[ext])
		{
			CXmlNodeWrapper node(root.InsertChildBefore(bnode.Interface(),"Default"));
			node.SetAttrValue("Extension",ext);
			node.SetAttrValue("ContentType",m_content_types[ext]);
			unique_exts[ext] = true;
		}
	}
	m_xml.Save();
}

void CDocxGenerator::Write(char *dstfile)
{
	string outfile = dstfile;
		
	// Get temporary path
	char s[MAX_PATH];
	GetTempPath(MAX_PATH,s);
	string path = s;
	path = path + "\\"+"img2docx";

	// Check whether \img2docx directory exists and create if not
	CreateDirectory(path.c_str(), NULL);

	// Clear the directory
	ClearDirectory(path);

	// Read template into buffer
	ReadTemplate();

	// Unzip template
	UnzipTemplate(path);
	
	// Copy images
	CopyImages(path);

	// Make XML changes
	MakeXMLChanges(path);
	
	// Zip it
	ZipTemplate(path);

	// Copy to target path
	CopyFile((path + "\\template.docx").c_str(),dstfile, false);
}