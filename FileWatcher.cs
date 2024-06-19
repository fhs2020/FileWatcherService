using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.ServiceProcess;
using OfficeOpenXml;

public partial class FileWatcher : ServiceBase
{
    private FileSystemWatcher watcher;

    public FileWatcher()
    {
        InitializeComponent();
    }

    private void InitializeComponent()
    {
        this.ServiceName = "FileWatcherService";
        watcher = new FileSystemWatcher
        {
            Path = @"C:\Home_File",
            Filter = "*.txt",
            EnableRaisingEvents = true
        };
        watcher.Created += OnCreated;
    }

    protected override void OnStart(string[] args)
    {
        watcher.EnableRaisingEvents = true;
    }

    protected override void OnStop()
    {
        watcher.EnableRaisingEvents = false;
    }

    private void OnCreated(object sender, FileSystemEventArgs e)
    {
        try
        {
            string excelPath = ConvertTextToExcel(e.FullPath);
            SendEmail(excelPath);
        }
        catch (Exception ex)
        {
            throw new Exception(ex.Message);
        }
    }

    private string ConvertTextToExcel(string filePath)
    {
        string excelPath = Path.ChangeExtension(filePath, ".xlsx");

        using (var package = new ExcelPackage())
        {
            var worksheet = package.Workbook.Worksheets.Add("Sheet1");
            var lines = File.ReadAllLines(filePath);

            for (int i = 0; i < lines.Length; i++)
            {
                var values = lines[i].Split(';');
                for (int j = 0; j < values.Length; j++)
                {
                    worksheet.Cells[i + 1, j + 1].Value = values[j];
                }
            }

            package.SaveAs(new FileInfo(excelPath));
        }

        return excelPath;
    }

    private void SendEmail(string attachmentPath)
    {
        string smtpServer = "smtp.office365.com";
        int smtpPort = 587;
        string fromEmail = "cloudgridnetworksolutions@outlook.com";
        string password = "baller100";
        string toEmail = "flavioh007@gmail.com";

        using (var client = new SmtpClient(smtpServer, smtpPort))
        {
            client.Credentials = new NetworkCredential(fromEmail, password);
            client.EnableSsl = true;

            var mail = new MailMessage(fromEmail, toEmail)
            {
                Subject = "New Excel File",
                Body = "Please find the attached Excel file."
            };

            mail.Attachments.Add(new Attachment(attachmentPath));

            try
            {
                client.Send(mail);
                Log("Email sent successfully.");
            }
            catch (SmtpException smtpEx)
            {
                Log($"SMTP Error: {smtpEx.Message}");
            }
            catch (Exception ex)
            {
                Log($"General Error: {ex.Message}");
            }
        }
    }

    private void Log(string message)
    {
        string logFilePath = @"C:\Home_File\service_log.txt";
        using (StreamWriter writer = new StreamWriter(logFilePath, true))
        {
            writer.WriteLine($"{DateTime.Now}: {message}");
        }
    }
}
