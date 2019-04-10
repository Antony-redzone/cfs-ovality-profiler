#include "EmbeddedFile.h"
#include <stdio.h>
#include <io.h>


EmbeddedFile::EmbeddedFile(void)
{
}

EmbeddedFile::~EmbeddedFile(void)
{
}

void EmbeddedFile::MoveFileData(char *_pvFileName, int _FromFilePosition, int _ToFilePosition)
{
	FILE *f;
	fpos_t a;
	long totalFileLength;
	char *fileData;
	long dataBufferSize;

	_FromFilePosition--;
	_ToFilePosition--;
	
	f = fopen(_pvFileName,"wb");
	totalFileLength = _filelength((int)f);
	dataBufferSize = totalFileLength - _FromFilePosition;
	
	fileData = new char[dataBufferSize];
	
	a = (long) _FromFilePosition;
	fsetpos( f,  &a);
	fread(fileData,1,dataBufferSize,f);
	
	a = (long) _ToFilePosition;
	fsetpos(f, &a);
	fwrite(fileData,1,dataBufferSize,f);

	fclose(f);
	delete[] fileData;
}