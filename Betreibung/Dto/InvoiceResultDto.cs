namespace Betreibung.Dto; 

public class InvoiceResultDto {
    public InvoiceResultDto(string date, string accessNumber, string patientName, string contractNumber, Dictionary<string, int> invoiceComponents, int total) {
        Date = date;
        AccessNumber = accessNumber;
        PatientName = patientName;
        ContractNumber = contractNumber;
        InvoiceComponents = invoiceComponents;
        Total = total;
    }

    public string Date { get; set; }

    public string AccessNumber { get; set; }

    public string PatientName { get; set; }

    public string ContractNumber { get; set; }

    public IDictionary<string, int> InvoiceComponents { get; set; }

    public int Total { get; set; }
}
