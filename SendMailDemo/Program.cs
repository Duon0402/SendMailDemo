using SendMailDemo.Services;
using static SendMailDemo.Services.ExcelService;

class Program
{
    static void Main()
    {
        var booking = new BookingIexModel
        {
            Id = Guid.NewGuid().ToString(),
            CustomerCode = "CUST12345",
            CountryCode = "US",
            ReferenceBillCode = "REF98765",
            ExpressDeliveryStaffCode = "EXP001",
            PickupCompany = "ABC Logistics",
            PickupName = "John Doe",
            PickupPhoneNumber = "+1 234-567-890",
            PickupEmail = "john.doe@abclogistics.com",
            PickupAddress = "123 Main St, New York, NY, USA",
            DeliveryCompany = "XYZ Corp",
            DeliveryName = "Jane Smith",
            DeliveryPhoneNumber = "+1 987-654-321",
            DeliveryEmail = "jane.smith@xyzcorp.com",
            DeliveryAddress = "456 Elm St, Los Angeles, CA, USA",
            ExporterCompany = "Global Exports Ltd.",
            ExporterName = "Mike Johnson",
            ExporterPhoneNumber = "+44 1234-567-890",
            ExporterEmail = "mike.johnson@globalexports.com",
            ExporterAddress = "789 Oak St, London, UK",
            ImporterCompany = "Imports USA Inc.",
            ImporterName = "Alice Brown",
            ImporterPhoneNumber = "+1 222-333-4444",
            ImporterEmail = "alice.brown@importsusa.com",
            ImporterAddress = "321 Maple Ave, Chicago, IL, USA",
            PickupDate = DateTime.Now.AddDays(2),
            NumberOfCarton = 10,
            Weight = 150.5m,
            Dimension = 500.0m,
            Condition = "Good",
            Urgent = true,
            TrackingNumber = "TRK123456789",
            FileAttachment = "invoice.pdf",
            Status = 1,
            CreatedTime = DateTime.Now,
            CreatedUser = "admin",
            UpdatedTime = DateTime.Now,
            UpdatedUser = "admin",
            CustomerCompany = "ABC Logistics",
            CustomerName = "John Doe"
        };
        ExcelService excelService = new ExcelService();
        var excelFile = excelService.PrepareExcelFile(booking);

        SendMailService mailService = new SendMailService();
        mailService.SendMail(excelFile);

        Console.ReadLine();
    }
}
