using System.Net;
using System.Net.Mail;

namespace SendMailDemo.Services
{
    public class SendMailService
    {
        public void SendMail(MemoryStream? fileStream = null, string fileName = "fileName.xlsx")
        {
            try
            {
                // Thông tin cấu hình SMTP
                string fromEmail = "duongdangtruong.it@gmail.com"; // Địa chỉ email gửi
                string password = "ilxf kfcx bvjz goaz"; // Mật khẩu email hoặc App Password
                string smtpServer = "smtp.gmail.com"; // SMTP server của Gmail
                int smtpPort = 587; // Cổng SMTP

                // Tạo đối tượng MailMessage
                MailMessage mail = new MailMessage();
                mail.From = new MailAddress(fromEmail);
                mail.To.Add("dangduong0402@gmail.com"); // Địa chỉ email nhận
                //mail.To.Add("duongdt@hncjsc.vn"); // Địa chỉ email nhận

                mail.Subject = "test"; // Tiêu đề email
                mail.Body = "Demo mail attachment excel file";
                mail.IsBodyHtml = true; // Cho phép HTML nếu cần

                if (fileStream != null)
                {
                    // Reset MemoryStream về đầu trước khi gửi
                    fileStream.Position = 0;

                    if (!string.IsNullOrEmpty(fileName) && !fileName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                    {
                        fileName += ".xlsx";
                    }

                    // Đính kèm file Excel
                    Attachment attachment = new Attachment(fileStream, fileName, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
                    mail.Attachments.Add(attachment);
                }

                // Cấu hình SMTP Client
                SmtpClient smtpClient = new SmtpClient(smtpServer, smtpPort);
                smtpClient.Credentials = new NetworkCredential(fromEmail, password);
                smtpClient.EnableSsl = true; // Sử dụng SSL để bảo mật

                // Gửi email
                smtpClient.Send(mail);

                Console.WriteLine("Email đã được gửi thành công.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Gửi email thất bại: " + ex.Message);
            }
        }
    }
}
