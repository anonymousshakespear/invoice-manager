namespace Betreibung.Dto;

using Betreibung.Enums;
using Betreibung.Helpers;

public class InvoiceInformationDto {
    public InvoiceInformationDto(string data) {
        var splits = data.Split(";");

        IsClosed = splits[0] == "Yes";
        Date = DateTime.Parse(splits[1]);
        AccessNumber = splits[2];
        PatientName = splits[3];
        ContractNumber = splits[4];
        CustomerCode = splits[5];
        Name = splits[6];
        Item = splits[7];
        DisplayName = splits[8];
        Amount = int.Parse(splits[9]);
        Area = EnumHelper.ParseEnum<InvoiceAreaEnum>(splits[10]);
        LegalName = splits[11];
        ShortName = splits[12];
        FileName = splits[13];

        Month = Date.Month;
    }

    public bool IsClosed { get; set; }

    public DateTime Date { get; set; }

    public string AccessNumber { get; set; }

    public string PatientName { get; set; }

    public string ContractNumber { get; set; }

    public string CustomerCode { get; set; }

    public string Name { get; set; }

    public string Item { get; set; }

    public string DisplayName { get; set; }
    
    public int Amount { get; set; }

    public InvoiceAreaEnum Area { get; set; }

    public string LegalName { get; set; }

    public string ShortName { get; set; }

    public string FileName { get; set; }

    public int Month { get; set; }
}
