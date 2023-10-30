namespace Betreibung.Dto; 

public class InvoiceResultDto {
    public InvoiceResultDto(DateTime date, string accessNumber, string patientName, string contractNumber, Dictionary<string, decimal> invoiceComponents, decimal total) {
        Date = date;
        AccessNumber = accessNumber;
        PatientName = patientName;
        ContractNumber = contractNumber;
        InvoiceComponents = invoiceComponents;
        Total = total;
    }

    public DateTime Date { get; set; }

    public string AccessNumber { get; set; }

    public string PatientName { get; set; }

    public string ContractNumber { get; set; }

    public IDictionary<string, decimal> InvoiceComponents { get; set; }

    public decimal Total { get; set; }
}
