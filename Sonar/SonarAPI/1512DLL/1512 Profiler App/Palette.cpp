#include "StdAfx.h"
#include ".\palette.h"

namespace Sonar
{

CPalette::CPalette(void)
: m_Loaded(false)
{
	ZeroMemory(&m_Color[0], sizeof(D3DCOLOR)*256);
}

CPalette::~CPalette(void)
{
}

void CPalette::Load(const char* filename)
{
	std::ifstream in;
	in.open(filename);
	
	if(in.is_open())
	{
		m_Loaded = true;

		char str[10];
		for (int i = 0; i < 256; i++)
		{
			char red[3];
			char green[3];
			char blue[3];

			red[2] = '\0';
			green[2] = '\0';
			blue[2] = '\0';

			in.getline(str, 10, '\n');

			red[0] = str[0];
			red[1] = str[1];

			green[0] = str[3];
			green[1] = str[4];
			
			blue[0] = str[6];
			blue[1] = str[7];
			
			m_Color[i] = RGB(atoi(red)*4, atoi(green)*4, atoi(blue)*4);
		};

		in.close();
		
	}
}

DWORD* CPalette::GetColor(int index)
{
	if(index >=0 && index < 256)
		return &m_Color[index];
	else return NULL;
}

bool CPalette::isLoaded(void)
{
	return m_Loaded;
}

};