namespace Betreibung.Processors;

using Betreibung.Dto;

public class CsvProcessor {
    public CsvProcessor() { }

    public void ProcessCsvDataToDto(List<string>? data) {
        if (data == null)
            return;

        var invoiceInformation = data.Select(x => new InvoiceInformationDto(x));
        var filteredInvoice = invoiceInformation.OrderBy(x => x.Name).GroupBy(x => x.Name);

        foreach (var companyInvoice in filteredInvoice) {
            var monthFilteredInvoice = companyInvoice.OrderBy(x => x.Month).GroupBy(x => x.Month);

            foreach (var monthlyInvoice in monthFilteredInvoice) {
                var invoiceList = new List<InvoiceResultDto>();
                var invoiceDisplays = new Dictionary<string, string>();
                var monthlyMainData = monthlyInvoice.First();
                var patientFilteredInvoice = monthlyInvoice.OrderBy(x => x.PatientName).GroupBy(x => x.PatientName);

                foreach (var patientInvoice in patientFilteredInvoice) {
                    var invoiceComponents = new Dictionary<string, int>();
                    var mainData = patientInvoice.First();
                    var total = 0;

                    foreach(var patient in patientInvoice) {
                        invoiceComponents.Add(patient.Item, patient.Amount);
                        total += patient.Amount;
                        if (!invoiceDisplays.ContainsKey(patient.Item)) {
                            invoiceDisplays.Add(patient.Item, patient.DisplayName);
                        }
                    }

                    invoiceList.Add(new InvoiceResultDto(mainData.Date.ToShortDateString(), mainData.AccessNumber, mainData.PatientName, mainData.ContractNumber, invoiceComponents, total));
                }

                var excelProcessor = new ExcelProcessor(invoiceList, invoiceDisplays, monthlyMainData.LegalName, monthlyMainData.FileName, monthlyMainData.Date);
                excelProcessor.ProcessExcelFile();
            }
        }
    }
}
