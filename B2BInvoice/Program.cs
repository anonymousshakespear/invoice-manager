namespace B2BInvoice;

using System.Configuration;
using Common.Handlers;
using B2BInvoice.Processors;

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