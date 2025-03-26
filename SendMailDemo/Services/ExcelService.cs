using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace SendMailDemo.Services
{
    public class ExcelService
    {
        public ExcelService()
        {
            // Thiết lập context license cho EPPlus (NonCommercial hoặc Commercial tùy theo mục đích sử dụng)
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        public MemoryStream PrepareExcelFile(string sheetName = "Sheet1", object? data = null)
        {
            var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add(sheetName);

            // Cột A
            SetCellStyle(worksheet, "A4", "DST OCS", false, ExcelBorderStyle.None, ExcelHorizontalAlignment.Right);

            SetCellStyle(worksheet, "A7", null, false, ExcelBorderStyle.Thin);
            SetCellStyle(worksheet, "A8", "*Company", true, ExcelBorderStyle.Thin, ExcelHorizontalAlignment.Right);
            SetCellStyle(worksheet, "A9", "*Attention Name", true, ExcelBorderStyle.Thin, ExcelHorizontalAlignment.Right);
            SetCellStyle(worksheet, "A10", "*Contact Phone No.", true, ExcelBorderStyle.Thin, ExcelHorizontalAlignment.Right);
            SetCellStyle(worksheet, "A11", "*Email address", true, ExcelBorderStyle.Thin, ExcelHorizontalAlignment.Right);
            SetCellStyle(worksheet, "A12", "*Address", true, ExcelBorderStyle.Thin, ExcelHorizontalAlignment.Right);

            SetCellStyle(worksheet, "A14", null, false, ExcelBorderStyle.Thin);
            SetCellStyle(worksheet, "A15", "Company", false, ExcelBorderStyle.Thin, ExcelHorizontalAlignment.Right);
            SetCellStyle(worksheet, "A16", "Attention Name", false, ExcelBorderStyle.Thin, ExcelHorizontalAlignment.Right);
            SetCellStyle(worksheet, "A17", "Contact Phone No.", false, ExcelBorderStyle.Thin, ExcelHorizontalAlignment.Right);
            SetCellStyle(worksheet, "A18", "*Email address", true, ExcelBorderStyle.Thin, ExcelHorizontalAlignment.Right);
            SetCellStyle(worksheet, "A19", "Address", false, ExcelBorderStyle.Thin, ExcelHorizontalAlignment.Right);

            SetCellStyle(worksheet, "A21", "*Goods Description", true, ExcelBorderStyle.Thin, ExcelHorizontalAlignment.Right);
            SetCellStyle(worksheet, "A22", "Pick Up Date", false, ExcelBorderStyle.Thin, ExcelHorizontalAlignment.Right);
            SetCellStyle(worksheet, "A23", "Number of Carton", false, ExcelBorderStyle.Thin, ExcelHorizontalAlignment.Right);
            SetCellStyle(worksheet, "A24", "Weight (kg)", false, ExcelBorderStyle.Thin, ExcelHorizontalAlignment.Right);
            SetCellStyle(worksheet, "A25", "Dimension (cm)", false, ExcelBorderStyle.Thin, ExcelHorizontalAlignment.Right);
            SetCellStyle(worksheet, "A26", "*Condition", true, ExcelBorderStyle.Thin, ExcelHorizontalAlignment.Right);

            SetCellValue(worksheet, "A27:G27", "Special Notation");
            SetCellValue(worksheet, "A28:G31", "If you attach INVOICE as a reference, please attach it with PDF.");

            // Cột B
            SetCellStyle(worksheet, "B1", "Request Form for IEX IMPORT", false, ExcelBorderStyle.None);
            SetCellStyle(worksheet, "B2", "Please fill in all required fields ( * ) in English only.", false, ExcelBorderStyle.None);
            SetCellStyle(worksheet, "B4:C4", "TEST DST_OCCS", true, ExcelBorderStyle.Thin, ExcelHorizontalAlignment.Center);

            SetCellStyle(worksheet, "B6:D6", "SHIPPER", true, ExcelBorderStyle.Thin, ExcelHorizontalAlignment.Center);
            SetCellStyle(worksheet, "B7:D7", "Pick Up Place", true, ExcelBorderStyle.Thin);
            SetCellStyle(worksheet, "B8:D8", "Company Shipper", false, ExcelBorderStyle.Thin);
            SetCellStyle(worksheet, "B9:D9", "Company Attention Name", false, ExcelBorderStyle.Thin);
            SetCellStyle(worksheet, "B10:D10", "Company Contact Phone No", false, ExcelBorderStyle.Thin);
            SetCellStyle(worksheet, "B11:D11", "Company Email Adress", false, ExcelBorderStyle.Thin);
            SetCellStyle(worksheet, "B12:D12", "Company Adress", false, ExcelBorderStyle.Thin);

            SetCellValue(worksheet, "B13:G13", "↓Please state if  Pick up place and Exporter / Delivery place and importer are different.");
            SetCellStyle(worksheet, "B14:D14", "Exporter on the Invoice", false, ExcelBorderStyle.Thin);
            SetCellStyle(worksheet, "B15:D15", "Exporter Shipper", false, ExcelBorderStyle.Thin);
            SetCellStyle(worksheet, "B16:D16", "Exporter Attention Name", false, ExcelBorderStyle.Thin);
            SetCellStyle(worksheet, "B17:D17", "Exporter Contact Phone No", false, ExcelBorderStyle.Thin);
            SetCellStyle(worksheet, "B18:D18", "Exporter Email Adress", false, ExcelBorderStyle.Thin);
            SetCellStyle(worksheet, "B19:D19", "Exporter Adress", false, ExcelBorderStyle.Thin);

            SetCellValue(worksheet, "B20:G20", "↓We need goods description for pick-up arrangement.\r\n Please write the goods name clearly. (Good) \"Bolt\" (Not Good) \"Machine Parts\"");
            SetCellStyle(worksheet, "B21:G21", "Goods Description", false, ExcelBorderStyle.Thin, ExcelHorizontalAlignment.Center);

            SetCellStyle(worksheet, "B22:G22", "26/03/2025", false, ExcelBorderStyle.Thin);
            SetCellStyle(worksheet, "B23:G23", "5", false, ExcelBorderStyle.Thin);
            SetCellStyle(worksheet, "B24:G24", "150 (kg)", false, ExcelBorderStyle.Thin);
            SetCellStyle(worksheet, "B25:G25", "500 (cm)", false, ExcelBorderStyle.Thin);
            SetCellStyle(worksheet, "B26:G26", "COMMERCIAL", false, ExcelBorderStyle.Thin);

            //Cột D
            SetCellValue(worksheet, "D4:E4", "Person In Charge");

            // Cột E
            SetCellStyle(worksheet, "E6:G6", "CONSIGNEE", true, ExcelBorderStyle.Thin, ExcelHorizontalAlignment.Center);

            SetCellStyle(worksheet, "E7:G7", "Delivery Place", true, ExcelBorderStyle.Thin);
            SetCellStyle(worksheet, "E8:G8", "Consignee Shipper", false, ExcelBorderStyle.Thin);
            SetCellStyle(worksheet, "E9:G9", "Consignee Attention Name", false, ExcelBorderStyle.Thin);
            SetCellStyle(worksheet, "E10:G10", "Consignee Contact Phone No", false, ExcelBorderStyle.Thin);
            SetCellStyle(worksheet, "E11:G11", "Consignee Email Adress", false, ExcelBorderStyle.Thin);
            SetCellStyle(worksheet, "E12:G12", "Consignee Adress", false, ExcelBorderStyle.Thin);

            SetCellStyle(worksheet, "E14:G14", "Importer Shipper", false, ExcelBorderStyle.Thin);
            SetCellStyle(worksheet, "E15:G15", "Importer Attention Name", false, ExcelBorderStyle.Thin);
            SetCellStyle(worksheet, "E16:G16", "Importer Contact Phone No", false, ExcelBorderStyle.Thin);
            SetCellStyle(worksheet, "E17:G17", "Importer Email Adress", false, ExcelBorderStyle.Thin);
            SetCellStyle(worksheet, "E18:G18", "Importer Adress", false, ExcelBorderStyle.Thin);
            SetCellStyle(worksheet, "E19:G19", "Importer Shipper", false, ExcelBorderStyle.Thin);


            // Cột F
            SetCellStyle(worksheet, "F1", "*World Account", true, ExcelBorderStyle.Thin);
            SetCellStyle(worksheet, "F4:G4", "TEST PersonInCharge", true, ExcelBorderStyle.Thin, ExcelHorizontalAlignment.Center);

            // Cột G
            SetCellStyle(worksheet, "G1", "TEST WorldAccount", false, ExcelBorderStyle.Thin);

            worksheet.Cells.AutoFitColumns();

            var stream = new MemoryStream();
            package.SaveAs(stream);
            stream.Position = 0;

            return stream;
        }

        private void SetCellStyle(ExcelWorksheet worksheet, string cellAddress, string? cellValue,
            bool bold = false,
            ExcelBorderStyle excelBorderStyle = ExcelBorderStyle.None,
            ExcelHorizontalAlignment horizontalAlignment = ExcelHorizontalAlignment.Left)
        {
            MergeCell(worksheet, cellAddress);
            var cell = worksheet.Cells[cellAddress];
            if (!string.IsNullOrWhiteSpace(cellValue))
            {
                cell.Value = cellValue;
            }

            cell.Style.Font.Bold = bold;

            cell.Style.Border.BorderAround(excelBorderStyle);

            cell.Style.HorizontalAlignment = horizontalAlignment;
        }

        private void SetTextStyle(ExcelWorksheet worksheet, string cellAddress, bool bold = false, bool italic = false, bool underline = false)
        {
            MergeCell(worksheet, cellAddress);
            var cell = worksheet.Cells[cellAddress];
            cell.Style.Font.Bold = bold;
            cell.Style.Font.Italic = italic;
            cell.Style.Font.UnderLine = underline;
        }

        private void SetCellValue(ExcelWorksheet worksheet, string cellAddress, string? cellValue)
        {
            MergeCell(worksheet, cellAddress);
            var cell = worksheet.Cells[cellAddress];
            if (!string.IsNullOrWhiteSpace(cellValue))
            {
                cell.Value = cellValue;
            }
        }

        private void MergeCell(ExcelWorksheet worksheet, string cellAddress)
        {
            if (!string.IsNullOrWhiteSpace(cellAddress) && cellAddress.Contains(":"))
            {
                worksheet.Cells[cellAddress].Merge = true;
            }
        }
    }
}
