#include "img2docx.h"
#include "docxgen.h"

CDocxGenerator g_docxgen;

void AddImage(char *scrfile)
{
	g_docxgen.Add(scrfile);
}

void CreateDocx(char *dstfile)
{
	g_docxgen.Write(dstfile);
}

void Reset(void)
{
	g_docxgen.Reset();
}