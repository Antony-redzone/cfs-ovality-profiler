class EmbeddedFile
{
public:
	EmbeddedFile(void);
	~EmbeddedFile(void);
	void MoveFileData(char *_FileName, int _FromFilePosition, int _ToFilePosition);
private:
};