namespace Betreibung.Processors;

using System.Configuration;
using System.Drawing;
using System.Runtime.InteropServices;
using Betreibung.Dto;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

public class ExcelProcessor {
    private readonly List<InvoiceResultDto> data;
    private readonly IDictionary<string, string> invoiceComponents;
    private readonly string legalName;
    private readonly string fileName;
    private readonly DateTime date;
    private int TotalLocation;
    private IDictionary<string, int> componentLocations;
    private IDictionary<string, string> componentTotals;

    public ExcelProcessor(List<InvoiceResultDto> data, IDictionary<string, string> invoiceComponents, string legalName, string fileName, DateTime date) {
        this.data = data;
        this.invoiceComponents = invoiceComponents;
        this.legalName = legalName;
        this.fileName = fileName;
        this.date = date;

        componentLocations = new Dictionary<string, int>();
        componentTotals = new Dictionary<string, string>();
    }

    public void ProcessExcelFile() {
        Application? excel = null;
        Workbooks? workBooks = null;
        Workbook? workBook = null;
        Worksheet? workSheet = null;

        try {
            excel = new Application();
            workBooks = excel.Workbooks;
            workBook = workBooks.Add();
            workSheet = (Worksheet)excel.ActiveSheet;

            excel.StandardFont = "Times New Roman";
            excel.StandardFontSize = 16;
            excel.ActiveWindow.Zoom = 55;

            DecorateExcelFile(workSheet);
            var rowCount = 13;
            foreach (var rowData in data) {
                workSheet.Cells[rowCount, "A"] = data.IndexOf(rowData) + 1;
                workSheet.Cells[rowCount, "B"] = $"'{rowData.Date.ToShortDateString()}";
                workSheet.Cells[rowCount, "C"] = rowData.AccessNumber;
                workSheet.Cells[rowCount, "D"] = rowData.PatientName;
                workSheet.Cells[rowCount, "E"] = rowData.ContractNumber;

                foreach (var components in rowData.InvoiceComponents) {
                    workSheet.Cells[rowCount, componentLocations[components.Key]] = components.Value;
                    if (!componentTotals.ContainsKey(components.Key))
                        componentTotals.Add(components.Key, $"{workSheet.Cells[rowCount, componentLocations[components.Key]].Address}:");
                    else
                        componentTotals[components.Key] += $"{workSheet.Cells[rowCount, componentLocations[components.Key]].Address}:";
                }

                workSheet.Cells[rowCount, TotalLocation].Formula = 
                    $"=SUM({workSheet.Cells[rowCount, "F"].Address}:{workSheet.Cells[rowCount, TotalLocation - 1].Address})";
                workSheet.Range[workSheet.Cells[rowCount, 1], workSheet.Cells[rowCount, TotalLocation]].Borders.Color = ColorTranslator.ToOle(Color.Black);

                workSheet.Cells[rowCount, TotalLocation + 2] = rowData.Total;
                workSheet.Cells[rowCount, TotalLocation + 3].Formula =
                    $"={workSheet.Cells[rowCount, TotalLocation + 2].Address}-{workSheet.Cells[rowCount, TotalLocation].Address}";
                rowCount++;
            }

            FinalizeInvoice(workSheet, rowCount);
            rowCount += 2;

            CoverLetter(workSheet, rowCount);

            workSheet.Range[workSheet.Rows[1], workSheet.Rows[TotalLocation + 2]].Columns.AutoFit();

            Console.WriteLine($"Saving {fileName}");
            workBook.SaveAs($"{ConfigurationManager.AppSettings.Get("outputDirectory")}\\{fileName}.xlsx", XlFileFormat.xlOpenXMLWorkbook);
        }
        finally {
            workBook.Close();
            excel.Quit();
            Release(workSheet);
            Release(workBook);
            Release(workBooks);
            Release(excel);
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }

    private void DecorateExcelFile(Worksheet workSheet) {
        var background = workSheet.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, workSheet.Cells[1, 1].Left, workSheet.Cells[1, 1].Top, 600, 100);
        background.Fill.ForeColor.RGB = Color.White.ToArgb();
        background.Line.Visible = MsoTriState.msoFalse;
        workSheet.Shapes.AddPicture(AppDomain.CurrentDomain.BaseDirectory + @"\Static\Images\diag.png", MsoTriState.msoFalse, MsoTriState.msoCTrue, workSheet.Cells[1, 1].Left, workSheet.Cells[1, 1].Top, 800, 100);

        workSheet.Cells[6, "E"] = $"Kính gửi: {legalName}";
        workSheet.Cells[7, "E"] = "Bảng Báo Cáo Thanh toán Chi Phí Y Tế";
        workSheet.Cells[8, "E"] = $"Tháng {date.Month}/{date.Year}";
        workSheet.Cells[9, "E"] = "ĐVT: 1.000 đ";
        workSheet.Cells[12, "A"] = "STT";
        workSheet.Cells[12, "B"] = "Date";
        workSheet.Cells[12, "C"] = "Access Number";
        workSheet.Cells[12, "D"] = "Patient Name";
        workSheet.Cells[12, "E"] = "HD";
        var currentCell = 6;
        foreach (var keyValuePair in invoiceComponents) {
            workSheet.Cells[12, currentCell] = keyValuePair.Value;
            componentLocations.Add(keyValuePair.Key, currentCell);
            currentCell++;
        }

        workSheet.Cells[12, currentCell] = "Grand total";
        TotalLocation = currentCell;

        workSheet.Range[workSheet.Cells[12, 1], workSheet.Cells[12, currentCell]].Borders.Color = ColorTranslator.ToOle(Color.Black);
        workSheet.Range[workSheet.Cells[12, 1], workSheet.Cells[12, currentCell]].Interior.Color = Color.FromArgb(217, 225, 242);

        workSheet.Cells[12, currentCell + 2] = "ERP Gốc";

        workSheet.Cells[12, currentCell + 3] = "DIFF";

        workSheet.Range[workSheet.Cells[6, "E"], workSheet.Cells[8, "E"]].Font.Bold = true;
        workSheet.Range[workSheet.Cells[6, "E"], workSheet.Cells[8, "E"]].Font.Name = "Times New Roman";
        workSheet.Range[workSheet.Cells[6, "E"], workSheet.Cells[8, "E"]].Font.Size = 16;

        workSheet.Rows[12].Font.Bold = true;
        workSheet.Rows[12].Font.Name = "Times New Roman";
        workSheet.Rows[12].Font.Size = 16;


        workSheet.Range[workSheet.Cells[6, "E"], workSheet.Cells[6, "J"]].Merge(true);
        workSheet.Range[workSheet.Cells[7, "E"], workSheet.Cells[7, "H"]].Merge(true);
    }

    private void FinalizeInvoice(Worksheet workSheet, int rowCount) {
        workSheet.Cells[rowCount, "A"] = "TOTAL";
        var result = string.Empty;
        foreach (var total in componentTotals) {
            workSheet.Cells[rowCount, componentLocations[total.Key]].Formula = $"=SUM({total.Value.Remove(total.Value.Length - 1)})";
            workSheet.Cells[rowCount, componentLocations[total.Key]].Borders.Color = ColorTranslator.ToOle(Color.Black);
            workSheet.Cells[rowCount, componentLocations[total.Key]].Interior.Color = Color.FromArgb(217, 225, 242);
            result += $"{workSheet.Cells[rowCount, componentLocations[total.Key]].Address}:";
        }

        workSheet.Cells[rowCount, TotalLocation].Formula = $"=Sum({result.Remove(result.Length - 1)})";
        workSheet.Cells[rowCount, TotalLocation].Borders.Color = ColorTranslator.ToOle(Color.Black);
        workSheet.Cells[rowCount, TotalLocation].Interior.Color = Color.FromArgb(217, 225, 242);

        workSheet.Cells[rowCount, TotalLocation + 2] = workSheet.Cells[rowCount, TotalLocation].Value;
        workSheet.Cells[rowCount, TotalLocation + 3].Formula = $"={workSheet.Cells[rowCount, TotalLocation + 2].Address}-{workSheet.Cells[rowCount, TotalLocation].Address}";

        workSheet.Cells[9, TotalLocation].Formula = $"={workSheet.Cells[rowCount, TotalLocation].Address}";
        workSheet.Cells[9, TotalLocation].Font.Bold = true;
        workSheet.Cells[9, TotalLocation].Font.Color = ColorTranslator.ToOle(Color.Red);
        workSheet.Cells[9, TotalLocation].Font.Underline = Microsoft.Office.Interop.Excel.XlUnderlineStyle.xlUnderlineStyleSingle;

        workSheet.Rows[rowCount].Font.Bold = true;
        workSheet.Range[workSheet.Cells[rowCount, 1], workSheet.Cells[rowCount, 5]].Borders.Color = ColorTranslator.ToOle(Color.Black);
        workSheet.Range[workSheet.Cells[rowCount, 1], workSheet.Cells[rowCount, 5]].Interior.Color = Color.FromArgb(217, 225, 242);
        workSheet.Range[workSheet.Cells[rowCount, 1], workSheet.Cells[rowCount, 5]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
        workSheet.Range[workSheet.Cells[rowCount, 1], workSheet.Cells[rowCount, 5]].Merge();
    }

    private void CoverLetter(Worksheet workSheet, int rowCount) {
        var formattedLocation = TotalLocation > 8 ? TotalLocation - 4 : 10;
        workSheet.Cells[rowCount, formattedLocation] = $"Tp.Hồ Chí Minh, ngày {DateTime.DaysInMonth(date.Year, date.Month)} tháng {date.Month} năm {date.Year}";
        workSheet.Cells[rowCount, formattedLocation].Font.Size = 20;
        workSheet.Cells[rowCount, formattedLocation].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
        workSheet.Range[workSheet.Cells[rowCount, formattedLocation], workSheet.Cells[rowCount, formattedLocation + 3]].Merge();

        workSheet.Cells[rowCount, formattedLocation + 6] = $"{date.Month}/{DateTime.DaysInMonth(date.Year, date.Month)}/{date.Year}";
        workSheet.Cells[rowCount, formattedLocation + 6].Font.Size = 18;
        workSheet.Cells[rowCount, formattedLocation + 6].Font.Bold = true;
        workSheet.Cells[rowCount, formattedLocation + 6].Interior.Color = Color.FromArgb(255, 255, 0);

        workSheet.Cells[rowCount, formattedLocation + 7] = "Ngày hóa đơn";
        workSheet.Cells[rowCount, formattedLocation + 7].Font.Size = 18;
        workSheet.Cells[rowCount, formattedLocation + 7].Font.Bold = true;

        rowCount++;

        workSheet.Cells[rowCount, 1] = "Người lập biểu";
        workSheet.Cells[rowCount, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
        workSheet.Cells[rowCount, 1].Font.Size = 20;
        workSheet.Range[workSheet.Cells[rowCount, 1], workSheet.Cells[rowCount, 3]].Merge();

        workSheet.Cells[rowCount, formattedLocation] = "Kế toán trưởng";
        workSheet.Cells[rowCount, formattedLocation].Font.Size = 20;
        workSheet.Cells[rowCount, formattedLocation].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
        workSheet.Range[workSheet.Cells[rowCount, formattedLocation], workSheet.Cells[rowCount, formattedLocation + 3]].Merge();

        rowCount += 7;

        workSheet.Cells[rowCount, 1] = "Nguyễn Quang Minh";
        workSheet.Cells[rowCount, 1].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
        workSheet.Cells[rowCount, 1].Font.Size = 20;
        workSheet.Cells[rowCount, 1].Font.Bold = true;
        workSheet.Range[workSheet.Cells[rowCount, 1], workSheet.Cells[rowCount, 3]].Merge();

        workSheet.Cells[rowCount, formattedLocation] = "Phan Hoàng Nguyên";
        workSheet.Cells[rowCount, formattedLocation].Font.Size = 20;
        workSheet.Cells[rowCount, formattedLocation].Font.Bold = true;
        workSheet.Cells[rowCount, formattedLocation].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
        workSheet.Range[workSheet.Cells[rowCount, formattedLocation], workSheet.Cells[rowCount, formattedLocation + 3]].Merge();
    }

    private void Release(object comObject) {
        Marshal.ReleaseComObject(comObject);
        comObject = null;
    }
}
