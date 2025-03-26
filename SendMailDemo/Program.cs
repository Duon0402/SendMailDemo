using SendMailDemo.Services;

class Program
{
    static void Main()
    {
        ExcelService excelService = new ExcelService();
        var excelFile = excelService.PrepareExcelFile();

        SendMailService mailService = new SendMailService();
        mailService.SendMail(excelFile);

        Console.ReadLine();
    }
}
