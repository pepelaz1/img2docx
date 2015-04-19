#include <windows.h>
#include "filehelper.h"

void ClearDirectory(string directory)
{
	string dir = directory + "\\*";
	WIN32_FIND_DATA ffdata;
	HANDLE find = FindFirstFile(dir.c_str(), &ffdata);
	do
	{
		if (strcmp(ffdata.cFileName,".") !=0 && 
			strcmp(ffdata.cFileName,"..") != 0)
		{
			if ( ffdata.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)
			{
				string cdir = directory + "\\" + ffdata.cFileName;
				ClearDirectory(cdir);
				RemoveDirectory(cdir.c_str());
			}
			else
			{
				string dfile = directory + "\\" + ffdata.cFileName;
				DeleteFile(dfile.c_str());
			}
		}
	}
	while(FindNextFileA(find, &ffdata));
	FindClose(find);
}

string GetFileExtension(string FileName)
{
    if(FileName.find_last_of(".") != std::string::npos)
        return FileName.substr(FileName.find_last_of(".")+1);
    return "";
}
