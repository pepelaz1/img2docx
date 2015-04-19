#include "imgloader.h"


CImgLoader::CImgLoader(void)
{
}


CImgLoader::~CImgLoader(void)
{
}

void CImgLoader::Reset(void)
{
	m_files.clear();
}

void CImgLoader::Add(char *file)
{
	m_files.push_back(file);
}
