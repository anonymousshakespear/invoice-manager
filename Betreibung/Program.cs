namespace Betreibung;

using System.Configuration;
using Betreibung.Handlers;
using Betreibung.Processors;

public class Program {
    static void Main(string[] args) {
        var csvFile = new FileHandler(ConfigurationManager.AppSettings.Get("directory") ?? string.Empty);
        csvFile.GetFileData();
        var data = csvFile.Data;
        data.RemoveAt(0);
        var processor = new CsvProcessor();
        processor.ProcessCsvDataToDto(data);
    }
}