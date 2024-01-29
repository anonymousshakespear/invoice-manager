namespace B2BInvoice.Processors;

using System.Configuration;
using System.Drawing;
using System.Runtime.InteropServices;
using B2BInvoice.Dto;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

public class ExcelProcessor {
    private readonly IList<BusinessInvoiceDto> data;
    private readonly BusinessInvoiceDto mainData;
    public ExcelProcessor(IList<BusinessInvoiceDto> data) {
        this.data = data;
        mainData = data.First();
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
            excel.StandardFontSize = 12;
            excel.ActiveWindow.Zoom = 80;

            DecorateExcelFile(workSheet);
            var rowCount = 12;
            foreach (var rowData in data) {
                workSheet.Cells[rowCount, "A"] = data.IndexOf(rowData) + 1;
                workSheet.Cells[rowCount, "B"] = $"'{rowData.Date.ToShortDateString()}";
                workSheet.Cells[rowCount, "C"] = rowData.AccessNumber;
                workSheet.Cells[rowCount, "D"] = rowData.PatientName;
                workSheet.Cells[rowCount, "E"] = rowData.DisplayName;
                workSheet.Cells[rowCount, "F"] = rowData.Amount;
                workSheet.Cells[rowCount, "G"] = rowData.Item;

                workSheet.Cells[rowCount, "F"].NumberFormat = "#,###";
                workSheet.Range[workSheet.Cells[rowCount, 1], workSheet.Cells[rowCount, 7]].Borders.Color = ColorTranslator.ToOle(Color.Black);

                rowCount++;
            }

            FinalizeInvoice(workSheet, rowCount);
            rowCount += 2;

            CoverLetter(workSheet, rowCount);

            workSheet.Range[workSheet.Rows[1], workSheet.Rows[9]].Columns.AutoFit();

            Console.WriteLine($"Saving {mainData.FileName}");
            workBook.SaveAs($"{ConfigurationManager.AppSettings.Get("outputDirectory")}\\{mainData.FileName}.xlsx", XlFileFormat.xlOpenXMLWorkbook);
        } finally {
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
        var background = workSheet.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, workSheet.Cells[1, 1].Left, workSheet.Cells[1, 1].Top, 400, 80);
        background.Fill.ForeColor.RGB = Color.White.ToArgb();
        background.Line.Visible = MsoTriState.msoFalse;
        workSheet.Shapes.AddPicture(AppDomain.CurrentDomain.BaseDirectory + @"\Static\Images\diag.png", MsoTriState.msoFalse, MsoTriState.msoCTrue, workSheet.Cells[1, 1].Left, workSheet.Cells[1, 1].Top, 600, 80);

        workSheet.Cells[7, "C"] = $"Kính gửi: {mainData.Name}";
        workSheet.Cells[8, "C"] = $"BẢNG TỔNG KẾT CHI TIẾT XÉT NGHIỆM THÁNG {mainData.Date.Month}/{mainData.Date.Year}";
        workSheet.Cells[9, "C"] = "ĐVT: VNĐ";
        workSheet.Cells[11, "A"] = "STT";
        workSheet.Cells[11, "B"] = "Date";
        workSheet.Cells[11, "C"] = "Access Number";
        workSheet.Cells[11, "D"] = "Patient Name";
        workSheet.Cells[11, "E"] = "Notes";
        workSheet.Cells[11, "F"] = "Test_Amt";
        workSheet.Cells[11, "G"] = "Item";

        workSheet.Range[workSheet.Cells[11, 1], workSheet.Cells[11, 7]].Borders.Color = ColorTranslator.ToOle(Color.Black);
        workSheet.Range[workSheet.Cells[11, 1], workSheet.Cells[11, 7]].Interior.Color = Color.FromArgb(217, 225, 242);

        workSheet.Range[workSheet.Cells[7, "C"], workSheet.Cells[8, "C"]].Font.Bold = true;
        workSheet.Range[workSheet.Cells[7, "C"], workSheet.Cells[8, "C"]].Font.Name = "Times New Roman";
        workSheet.Range[workSheet.Cells[7, "C"], workSheet.Cells[8, "C"]].Font.Size = 16;

        workSheet.Rows[11].Font.Bold = true;

        workSheet.Range[workSheet.Cells[7, "C"], workSheet.Cells[7, "F"]].Merge(true);
        workSheet.Range[workSheet.Cells[8, "C"], workSheet.Cells[8, "F"]].Merge(true);
    }

    private void FinalizeInvoice(Worksheet workSheet, int rowCount) {
        workSheet.Cells[rowCount, "A"] = "TOTAL";
        workSheet.Cells[rowCount, 6].Formula = $"=SUM({workSheet.Cells[12, 6].Address}:{workSheet.Cells[rowCount - 1, 6].Address})";
        workSheet.Cells[rowCount, 6].NumberFormat = "#,###";

        workSheet.Range[workSheet.Cells[rowCount, 1], workSheet.Cells[rowCount, 7]].Font.Bold = true;
        workSheet.Range[workSheet.Cells[rowCount, 1], workSheet.Cells[rowCount, 7]].Borders.Color = ColorTranslator.ToOle(Color.Black);
        workSheet.Range[workSheet.Cells[rowCount, 1], workSheet.Cells[rowCount, 7]].Interior.Color = Color.FromArgb(217, 225, 242);
        workSheet.Range[workSheet.Cells[rowCount, 1], workSheet.Cells[rowCount, 7]].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
        workSheet.Range[workSheet.Cells[rowCount, 1], workSheet.Cells[rowCount, 5]].Merge();

        workSheet.Cells[10, 6].Formula = $"={workSheet.Cells[rowCount, 6].Address}";
        workSheet.Cells[10, 6].Font.Bold = true;
        workSheet.Cells[10, 6].Font.Color = ColorTranslator.ToOle(Color.Red);
        workSheet.Cells[10, 6].Font.Underline = Microsoft.Office.Interop.Excel.XlUnderlineStyle.xlUnderlineStyleSingle;
    }

    private void CoverLetter(Worksheet workSheet, int rowCount) {
        workSheet.Cells[rowCount, 5] = $"Tp.Hồ Chí Minh, ngày {DateTime.DaysInMonth(mainData.Date.Year, mainData.Date.Month)} tháng {mainData.Date.Month} năm {mainData.Date.Year}";
        workSheet.Cells[rowCount, 5].Font.Size = 15;
        workSheet.Cells[rowCount, 5].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
        workSheet.Range[workSheet.Cells[rowCount, 5], workSheet.Cells[rowCount, 6]].Merge();

        workSheet.Cells[rowCount, 8] = $"{mainData.Date.Month}/{DateTime.DaysInMonth(mainData.Date.Year, mainData.Date.Month)}/{mainData.Date.Year}";
        workSheet.Cells[rowCount, 8].Font.Size = 15;
        workSheet.Cells[rowCount, 8].Font.Bold = true;
        workSheet.Cells[rowCount, 8].Interior.Color = Color.FromArgb(255, 255, 0);

        workSheet.Cells[rowCount, 9] = "Ngày hóa đơn";
        workSheet.Cells[rowCount, 9].Font.Size = 15;
        workSheet.Cells[rowCount, 9].Font.Bold = true;

        rowCount++;

        workSheet.Cells[rowCount, 2] = "Người lập biểu";
        workSheet.Cells[rowCount, 2].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
        workSheet.Cells[rowCount, 2].Font.Size = 15;
        workSheet.Range[workSheet.Cells[rowCount, 1], workSheet.Cells[rowCount, 3]].Merge();

        workSheet.Cells[rowCount, 5] = "Kế toán trưởng";
        workSheet.Cells[rowCount, 5].Font.Size = 15;
        workSheet.Cells[rowCount, 5].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
        workSheet.Range[workSheet.Cells[rowCount, 5], workSheet.Cells[rowCount, 6]].Merge();

        rowCount += 6;

        workSheet.Cells[rowCount, 2] = "Nguyễn Quang Minh";
        workSheet.Cells[rowCount, 2].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
        workSheet.Cells[rowCount, 2].Font.Size = 15;
        workSheet.Cells[rowCount, 2].Font.Bold = true;
        workSheet.Range[workSheet.Cells[rowCount, 1], workSheet.Cells[rowCount, 3]].Merge();

        workSheet.Cells[rowCount, 5] = "Chu Thị Lan Anh";
        workSheet.Cells[rowCount, 5].Font.Size = 15;
        workSheet.Cells[rowCount, 5].Font.Bold = true;
        workSheet.Cells[rowCount, 5].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
        workSheet.Range[workSheet.Cells[rowCount, 5], workSheet.Cells[rowCount, 6]].Merge();
    }

    private void Release(object comObject) {
        Marshal.ReleaseComObject(comObject);
        comObject = null;
    }
}
