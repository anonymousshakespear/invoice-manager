namespace Betreibung.Handlers; 

public class FileHandler {
    private readonly string fileName;

    public List<string> Data { get; set; } = new List<string>();

    public FileHandler(string fileName) {
        this.fileName = fileName;
    }

    public void GetFileData()
        => Data.AddRange(File.ReadLines(fileName));
}
