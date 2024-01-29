namespace B2BInvoice.Processors;

using B2BInvoice.Dto;

public class CsvProcessor {
    public void ProcessCsvDataToDto(List<string>? data) {
        if (data == null)
            return;

        var invoiceInformation = data.Select(x => new BusinessInvoiceDto(x));
        var filteredInvoice = invoiceInformation.OrderBy(x => x.CustomerCode).GroupBy(x => x.CustomerCode);

        foreach (var companyInvoice in filteredInvoice) {
            var monthFilteredInvoice = companyInvoice.OrderBy(x => x.Month).GroupBy(x => x.Month);

            foreach (var monthlyInvoice in monthFilteredInvoice) {
                var dateFilteredInvoice = monthlyInvoice.OrderBy(x => x.Date).ThenBy(x => x.AccessNumber);
                var excelProcessor = new ExcelProcessor(dateFilteredInvoice.ToList());
                excelProcessor.ProcessExcelFile();
            }
        }
    }
}
