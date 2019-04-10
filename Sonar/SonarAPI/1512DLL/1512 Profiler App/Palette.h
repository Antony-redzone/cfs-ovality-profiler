#pragma once
#include <d3d9.h>
#include <d3dx9.h>
#include <fstream>

namespace Sonar
{

	class CPalette
	{
		DWORD m_Color[256];
		bool m_Loaded;

	public:
		CPalette(void);
		virtual ~CPalette(void);

		void Load(const char* filename);
		DWORD *GetColor(int index);
		bool isLoaded(void);
	};

};
