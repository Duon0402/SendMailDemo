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

        public MemoryStream PrepareExcelFile(BookingIexModel? data = null, string sheetName = "Sheet1")
        {
            var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add(sheetName);

            // Cột A
            SetCellStyle(worksheet, "A3", "DST OCS", false, ExcelBorderStyle.None, ExcelHorizontalAlignment.Right);

            SetCellStyle(worksheet, "A6", null, false, ExcelBorderStyle.Thin);
            SetCellStyle(worksheet, "A7", "*Company", true, ExcelBorderStyle.Thin, ExcelHorizontalAlignment.Right);
            SetCellStyle(worksheet, "A8", "*Attention Name", true, ExcelBorderStyle.Thin, ExcelHorizontalAlignment.Right);
            SetCellStyle(worksheet, "A9", "*Contact Phone No.", true, ExcelBorderStyle.Thin, ExcelHorizontalAlignment.Right);
            SetCellStyle(worksheet, "A10", "*Email address", true, ExcelBorderStyle.Thin, ExcelHorizontalAlignment.Right);
            SetCellStyle(worksheet, "A11", "*Address", true, ExcelBorderStyle.Thin, ExcelHorizontalAlignment.Right);

            SetCellStyle(worksheet, "A13", null, false, ExcelBorderStyle.Thin);
            SetCellStyle(worksheet, "A14", "Company", false, ExcelBorderStyle.Thin, ExcelHorizontalAlignment.Right);
            SetCellStyle(worksheet, "A15", "Attention Name", false, ExcelBorderStyle.Thin, ExcelHorizontalAlignment.Right);
            SetCellStyle(worksheet, "A16", "Contact Phone No.", false, ExcelBorderStyle.Thin, ExcelHorizontalAlignment.Right);
            SetCellStyle(worksheet, "A17", "*Email address", true, ExcelBorderStyle.Thin, ExcelHorizontalAlignment.Right);
            SetCellStyle(worksheet, "A18", "Address", false, ExcelBorderStyle.Thin, ExcelHorizontalAlignment.Right);

            SetCellStyle(worksheet, "A20", "*Goods Description", true, ExcelBorderStyle.Thin, ExcelHorizontalAlignment.Right);
            SetCellStyle(worksheet, "A21", "Pick Up Date", false, ExcelBorderStyle.Thin, ExcelHorizontalAlignment.Right);
            SetCellStyle(worksheet, "A22", "Number of Carton", false, ExcelBorderStyle.Thin, ExcelHorizontalAlignment.Right);
            SetCellStyle(worksheet, "A23", "Weight (kg)", false, ExcelBorderStyle.Thin, ExcelHorizontalAlignment.Right);
            SetCellStyle(worksheet, "A24", "Dimension (cm)", false, ExcelBorderStyle.Thin, ExcelHorizontalAlignment.Right);
            SetCellStyle(worksheet, "A25", "*Condition", true, ExcelBorderStyle.Thin, ExcelHorizontalAlignment.Right);

            // Cột B
            SetCellStyle(worksheet, "B1", "Request Form for IEX IMPORT", false, ExcelBorderStyle.None);
            SetCellStyle(worksheet, "B3:C3", "TEST DST_OCCS", true, ExcelBorderStyle.Thin, ExcelHorizontalAlignment.Center);

            SetCellStyle(worksheet, "B5:D5", "SHIPPER", true, ExcelBorderStyle.Thin, ExcelHorizontalAlignment.Center);
            SetCellStyle(worksheet, "B6:D6", "Pick Up Place", true, ExcelBorderStyle.Thin);
            SetCellStyle(worksheet, "B7:D7", data.PickupCompany, false, ExcelBorderStyle.Thin);
            SetCellStyle(worksheet, "B8:D8", data.PickupName, false, ExcelBorderStyle.Thin);
            SetCellStyle(worksheet, "B9:D9", data.PickupPhoneNumber, false, ExcelBorderStyle.Thin);
            SetCellStyle(worksheet, "B10:D10", data.PickupEmail, false, ExcelBorderStyle.Thin);
            SetCellStyle(worksheet, "B11:D11", data.PickupAddress, false, ExcelBorderStyle.Thin);

            SetCellStyle(worksheet, "B13:D13", "Exporter on the Invoice", false, ExcelBorderStyle.Thin);
            SetCellStyle(worksheet, "B14:D14", data.ExporterCompany, false, ExcelBorderStyle.Thin);
            SetCellStyle(worksheet, "B15:D15", data.ExporterName, false, ExcelBorderStyle.Thin);
            SetCellStyle(worksheet, "B16:D16", data.ExporterPhoneNumber, false, ExcelBorderStyle.Thin);
            SetCellStyle(worksheet, "B17:D17", data.ExporterEmail, false, ExcelBorderStyle.Thin);
            SetCellStyle(worksheet, "B18:D18", data.ExporterAddress, false, ExcelBorderStyle.Thin);

            SetCellStyle(worksheet, "B20:G20", "Goods Description", false, ExcelBorderStyle.Thin, ExcelHorizontalAlignment.Center);

            SetCellStyle(worksheet, "B21:G21", data.PickupDate.ToString(), false, ExcelBorderStyle.Thin);
            SetCellStyle(worksheet, "B22:G22", data.NumberOfCarton.ToString(), false, ExcelBorderStyle.Thin);
            SetCellStyle(worksheet, "B23:G23", data.Weight.ToString(), false, ExcelBorderStyle.Thin);
            SetCellStyle(worksheet, "B24:G24", data.Dimension.ToString(), false, ExcelBorderStyle.Thin);
            SetCellStyle(worksheet, "B25:G25", data.Condition, false, ExcelBorderStyle.Thin);


            //Cột D
            SetCellValue(worksheet, "D3:E3", "Person In Charge");

            // Cột E
            SetCellStyle(worksheet, "E5:G5", "CONSIGNEE", true, ExcelBorderStyle.Thin, ExcelHorizontalAlignment.Center);

            SetCellStyle(worksheet, "E6:G6", "Delivery Place", true, ExcelBorderStyle.Thin);
            SetCellStyle(worksheet, "E7:G7", data.DeliveryCompany, false, ExcelBorderStyle.Thin);
            SetCellStyle(worksheet, "E8:G8", data.DeliveryName, false, ExcelBorderStyle.Thin);
            SetCellStyle(worksheet, "E9:G9", data.DeliveryPhoneNumber, false, ExcelBorderStyle.Thin);
            SetCellStyle(worksheet, "E10:G10", data.DeliveryEmail, false, ExcelBorderStyle.Thin);
            SetCellStyle(worksheet, "E11:G11", data.DeliveryAddress, false, ExcelBorderStyle.Thin);

            SetCellStyle(worksheet, "E13:G13", "Importer Shipper", false, ExcelBorderStyle.Thin);
            SetCellStyle(worksheet, "E14:G14", data.ImporterCompany, false, ExcelBorderStyle.Thin);
            SetCellStyle(worksheet, "E15:G15", data.ImporterName, false, ExcelBorderStyle.Thin);
            SetCellStyle(worksheet, "E16:G16", data.ImporterPhoneNumber, false, ExcelBorderStyle.Thin);
            SetCellStyle(worksheet, "E17:G17", data.ImporterEmail, false, ExcelBorderStyle.Thin);
            SetCellStyle(worksheet, "E18:G18", data.ImporterAddress, false, ExcelBorderStyle.Thin);


            // Cột F
            SetCellStyle(worksheet, "F1", "*World Account", true, ExcelBorderStyle.Thin);
            SetCellStyle(worksheet, "F3:G3", "TEST PersonInCharge", true, ExcelBorderStyle.Thin, ExcelHorizontalAlignment.Center);

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

        public class BookingIexModel
        {
            public string Id { get; set; }

            public string CustomerCode { get; set; }

            public string CountryCode { get; set; }

            public string ReferenceBillCode { get; set; }

            public string ExpressDeliveryStaffCode { get; set; }

            public string PickupCompany { get; set; }

            public string PickupName { get; set; }

            public string PickupPhoneNumber { get; set; }

            public string PickupEmail { get; set; }

            public string PickupAddress { get; set; }

            public string DeliveryCompany { get; set; }

            public string DeliveryName { get; set; }

            public string DeliveryPhoneNumber { get; set; }

            public string DeliveryEmail { get; set; }

            public string DeliveryAddress { get; set; }

            public string ExporterCompany { get; set; }

            public string ExporterName { get; set; }

            public string ExporterPhoneNumber { get; set; }

            public string ExporterEmail { get; set; }

            public string ExporterAddress { get; set; }

            public string ImporterCompany { get; set; }

            public string ImporterName { get; set; }

            public string ImporterPhoneNumber { get; set; }

            public string ImporterEmail { get; set; }

            public string ImporterAddress { get; set; }

            public DateTime PickupDate { get; set; }

            public decimal NumberOfCarton { get; set; }

            public decimal Weight { get; set; }

            public decimal Dimension { get; set; }

            public string Condition { get; set; }

            public bool Urgent { get; set; }

            public string TrackingNumber { get; set; }

            public string FileAttachment { get; set; }

            public int Status { get; set; }

            public DateTime CreatedTime { get; set; }

            public string CreatedUser { get; set; }

            public DateTime UpdatedTime { get; set; }

            public string UpdatedUser { get; set; }

            // trường không thuộc db
            public string CustomerCompany { get; set; }
            public string CustomerName { get; set; }
        }
    }
}
