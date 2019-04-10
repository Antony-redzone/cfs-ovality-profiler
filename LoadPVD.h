class LoadPVD
{
public:
	LoadPVD(char *_pvFileName,
				int _pvDataStartAddress,
				int _pvDataBlockSize,
				int _xy,
				float *_pvDataX,
                float *_pvDataY,
                double _pvDataMultiplier,
                int _fromFrame,
                int _toFrame);
	~LoadPVD(void);
	void LoadPVDData(void);
private:
	char *pvFileName;
	int pvDataStartAddress;
	int pvDataBlockSize;
	float *pvDataX;
    float *pvDataY;
    double pvDataMultiplier;
    int fromFrame;
    int toFrame;
	int xy_data; //If its radius this will be false, if XY then true
};

