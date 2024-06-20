using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.ServiceProcess;
using OfficeOpenXml;
using System.Configuration;

public partial class FileWatcher : ServiceBase
{
    private FileSystemWatcher watcher;
    private string watchFolderPath;
    private string logFolderPath;

    public FileWatcher()
    {
        InitializeComponent();
    }

    private void InitializeComponent()
    {
        this.ServiceName = "FileWatcherService";
        watchFolderPath = ConfigurationManager.AppSettings["WatchFolderPath"];
        logFolderPath = ConfigurationManager.AppSettings["LogFolderPath"];

        // Ensure the log directory exists
        if (!Directory.Exists(logFolderPath))
        {
            Directory.CreateDirectory(logFolderPath);
        }

        watcher = new FileSystemWatcher
        {
            Path = watchFolderPath,
            Filter = "*.txt",
            EnableRaisingEvents = true
        };
        watcher.Created += OnCreated;
    }

    protected override void OnStart(string[] args)
    {
        watcher.EnableRaisingEvents = true;
        Log("Service started.");
    }

    protected override void OnStop()
    {
        watcher.EnableRaisingEvents = false;
        Log("Service stopped.");
    }

    private void OnCreated(object sender, FileSystemEventArgs e)
    {
        try
        {
            Log($"File created: {e.FullPath}");
            string excelPath = ConvertTextToExcel(e.FullPath);
            Log($"Converted to Excel: {excelPath}");
            SendEmail(excelPath);
            Log("Email sent successfully.");
        }
        catch (Exception ex)
        {
            Log($"Error: {ex.Message}");
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
        string fromEmail = ConfigurationManager.AppSettings["FromEmail"];
        string password = ConfigurationManager.AppSettings["Password"];
        string toEmail = ConfigurationManager.AppSettings["ToEmail"];
        string smtpServer = ConfigurationManager.AppSettings["SmtpServer"];
        int smtpPort = int.Parse(ConfigurationManager.AppSettings["SmtpPort"]);

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
        string logFilePath = Path.Combine(logFolderPath, "service_log.txt");
        using (StreamWriter writer = new StreamWriter(logFilePath, true))
        {
            writer.WriteLine($"{DateTime.Now}: {message}");
        }
    }
}
