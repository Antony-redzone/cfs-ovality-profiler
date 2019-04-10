class SmoothGraph
{
public:
	SmoothGraph(float *_graphData, int _numberFrames);
	void Smooth(void);

private:
	float *graphData;
	float *tempData;
	int numberFrames;
};