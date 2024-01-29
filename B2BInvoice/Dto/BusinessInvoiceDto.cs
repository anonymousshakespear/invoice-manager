namespace B2BInvoice.Dto;

using System.Globalization;

public class BusinessInvoiceDto {
    public BusinessInvoiceDto(string data) {
        var splits = data.Split(";");

        Date = DateTime.ParseExact(splits[0], "dd/MM/yyyy", CultureInfo.InvariantCulture);
        AccessNumber = splits[1];
        PatientName = splits[2];
        Name = splits[3];
        CustomerCode = splits[4];
        DisplayName = splits[5];
        Item = splits[6];
        Amount = decimal.Parse(splits[7]);
        FileName = splits[8];

        Month = Date.Month;
    }

    public DateTime Date { get; set; }

    public string AccessNumber { get; set; }

    public string PatientName { get; set; }

    public string CustomerCode { get; set; }

    public string Name { get; set; }

    public string Item { get; set; }

    public string DisplayName { get; set; }

    public decimal Amount { get; set; }

    public string FileName { get; set; }

    public int Month { get; set; }
}
