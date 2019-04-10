#include <stdio.h>

class character {
public:
	character(void){
		height=0;
		width=0;
		data=NULL;
	}
	void copy(int h,int w,int *d){
		height=h;
		width=w;
		data = new int[h];
		for(int i=0;i<h;i++) *(data+i) = *(d+i);
	}
private:
	int height;
	int width;
	int *data;

};

class number_set {
public:
	number_set(void){
		//nothing here yet
	}

private:
	character set[10];

};

class character_database{

private:
	number_set *database;

};