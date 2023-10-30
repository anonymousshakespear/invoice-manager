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
                var invoiceLDisplays = new Dictionary<string, string>();
                var invoiceMDisplays = new Dictionary<string, string>();
                var monthlyMainData = monthlyInvoice.First();
                var patientFilteredInvoice = monthlyInvoice.OrderBy(x => x.AccessNumber).GroupBy(x => x.AccessNumber);

                foreach (var patientInvoice in patientFilteredInvoice) {
                    var invoiceComponents = new Dictionary<string, decimal>();
                    var mainData = patientInvoice.First();
                    decimal total = 0;

                    foreach(var patient in patientInvoice) {
                        invoiceComponents.Add(patient.Item, patient.Amount);
                        total += patient.Amount;
                        if (patient.Area == Enums.InvoiceAreaEnum.L && !invoiceLDisplays.ContainsKey(patient.Item))
                            invoiceLDisplays.Add(patient.Item, patient.DisplayName);
                        else if (patient.Area == Enums.InvoiceAreaEnum.M && !invoiceMDisplays.ContainsKey(patient.Item))
                            invoiceMDisplays.Add(patient.Item, patient.DisplayName);
                    }

                    invoiceList.Add(new InvoiceResultDto(mainData.Date, mainData.AccessNumber, mainData.PatientName, mainData.ContractNumber, invoiceComponents, total));
                }

                var invoiceDisplays = invoiceLDisplays.Concat(invoiceMDisplays)
                    .ToLookup(x => x.Key, x => x.Value).ToDictionary(x => x.Key, g => g.First());
                var excelProcessor = new ExcelProcessor(invoiceList, invoiceDisplays, monthlyMainData.LegalName, monthlyMainData.FileName, monthlyMainData.Date);
                excelProcessor.ProcessExcelFile();
            }
        }
    }
}
