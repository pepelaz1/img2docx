// testapp.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"
#include "windows.h"
#include "..\img2docx\img2docx.h"


#include <fstream>
using namespace std;

/*void MakeTemplate(void)
{
  int csize = 1000;
  unsigned char * buffer;
  long size;
  ifstream ifile ("template.docx", ios::in|ios::binary|ios::ate);
  size = (long)ifile.tellg();
  ifile.seekg (0, ios::beg);
  buffer = new unsigned char [size];
  ifile.read ((char *)buffer, size);
  ifile.close();
  ofstream ofile ("template.h", ios::out);
  ofile << "const char *g_template["<< size / csize+1 <<"]={";
  int j = 0;
  for(int i = 0; i < size; i++)
  {
	  if (!(i%csize))
	  {
		  if (i)
			ofile << "\"";

		  if (i + csize < size*2 && i)
			ofile << ",\"";
		  else
			ofile << "\"";
	  }
	  char sym[5];
	  sprintf_s(sym,"%02x",buffer[i]);
	  ofile << sym;
  }
  ofile << "\"};";
  ofile.close();
  delete[] buffer;
}*/

int _tmain(int argc, _TCHAR* argv[])
{
  //MakeTemplate();
  if (argc > 2)
  {
	CoInitialize(NULL);
	Reset();

	// Parse command line arguments
	for (int i = 1; i < argc-1; i++)
		AddImage(argv[i]);

	CreateDocx(argv[argc-1]);		
	CoUninitialize();
   }
   return 0;
}

