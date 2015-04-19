#pragma once
#include "xmlwrapper.h"
#include <list>
#include <string>
#include <map>
using namespace std;

#define TEMPLATE_SIZE  12675

struct WordImage
{
	string id;
	string ext;
};

class CDocxGenerator
{
private:
	// List of source filenames
	list<string> m_srcfiles;

	// Object for working with XML
	CXmlDocWrapper m_xml;

	// Map object for storing correspondance between image file extensions and their content types
	map<string,string> m_content_types;
	// Populating of content types
	void PopulateContentTypes();


	// Template buffer
	char m_tbuff[TEMPLATE_SIZE];

	// Read template
	void ReadTemplate(void);

	// Unzip template
	void UnzipTemplate(string path);

	// Zip template
	void ZipTemplate(string path);

	// Copy images
	void CopyImages(string path);

	// Make XML changes
	void MakeXMLChanges(string path);
	void AddContentType(string path);
public:
	CDocxGenerator(void);
	~CDocxGenerator(void);

	// Reset 
	void Reset(void);

	// Adds src image file into list
	void Add(char *srcfile);

	// Generate docx file
	void Write(char *dstfile);
};

