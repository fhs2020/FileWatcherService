using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.ServiceProcess;
using System.Configuration;
using iTextSharp.text;
using iTextSharp.text.pdf;

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
            string pdfPath = ConvertTextToPDF(e.FullPath);
            Log($"Converted to PDF: {pdfPath}");
            SendEmail(pdfPath);
            Log("Email sent successfully.");
        }
        catch (Exception ex)
        {
            Log($"Error: {ex.Message}");
        }
    }

    private string ConvertTextToPDF(string filePath)
    {
        string pdfPath = Path.ChangeExtension(filePath, ".pdf");

        using (var stream = new FileStream(pdfPath, FileMode.Create, FileAccess.Write, FileShare.None))
        {
            Document document = new Document();
            PdfWriter writer = PdfWriter.GetInstance(document, stream);
            document.Open();

            // Add the logos and title text
            AddLogosAndTitle(document);

            // Add fixed text
            AddFixedText(document);

            // Add dynamic text
            AddDynamicText(document, filePath);

            document.Close();
            writer.Close();
        }

        return pdfPath;
    }

    private void AddLogosAndTitle(Document document)
    {
        string logoPathLeft = ConfigurationManager.AppSettings["LogoPathLeft"];
        string logoPathRight = ConfigurationManager.AppSettings["LogoPathRight"];

        PdfPTable table = new PdfPTable(2);
        table.WidthPercentage = 100;

        // Add left logo and title text in the same cell
        if (File.Exists(logoPathLeft))
        {
            Image logoLeft = Image.GetInstance(logoPathLeft);
            logoLeft.ScaleToFit(150f, 150f);

            // Create a Phrase to combine logo and title text
            Phrase leftContent = new Phrase();
            leftContent.Add(new Chunk(logoLeft, 0, 0, true));
            leftContent.Add(new Chunk("\nOperações Pós-Vendas", new Font(Font.FontFamily.HELVETICA, 16)));

            PdfPCell cellLeft = new PdfPCell(leftContent)
            {
                Border = Rectangle.NO_BORDER,
                HorizontalAlignment = Element.ALIGN_LEFT,
                PaddingBottom = 0 // No space between logo and text
            };
            table.AddCell(cellLeft);
        }
        else
        {
            table.AddCell(new PdfPCell() { Border = Rectangle.NO_BORDER });
        }

        // Add right logo
        if (File.Exists(logoPathRight))
        {
            Image logoRight = Image.GetInstance(logoPathRight);
            logoRight.ScaleToFit(40f, 40f); // Changed size to be smaller
            PdfPCell cellRight = new PdfPCell(logoRight) { Border = Rectangle.NO_BORDER, HorizontalAlignment = Element.ALIGN_RIGHT };
            table.AddCell(cellRight);
        }
        else
        {
            table.AddCell(new PdfPCell() { Border = Rectangle.NO_BORDER });
        }

        document.Add(table);
    }

    private void AddFixedText(Document document)
    {
        Font smallerFont = new Font(Font.FontFamily.HELVETICA, 10); // Smaller font size

        // Add ATE-020/23 text with space above it
        document.Add(new Paragraph("\n\nATE-020/23\n\n", smallerFont));

        // Create a table to hold the city and date on the right below ATE-020/23
        PdfPTable table = new PdfPTable(1);
        table.WidthPercentage = 100;

        // Add the city and date to the table cell and align it to the right
        PdfPCell cell = new PdfPCell(new Phrase("São Bernardo do Campo, 28 de Febrero de 2023.\n\n", smallerFont));
        cell.Border = Rectangle.NO_BORDER;
        cell.HorizontalAlignment = Element.ALIGN_RIGHT;
        table.AddCell(cell);

        document.Add(table);

        // Add remaining fixed text
        document.Add(new Paragraph("0603 – HANSA LTDA.\n\n", smallerFont));
        document.Add(new Paragraph("LA PAZ - BOLIVIA\n\n", smallerFont));
        document.Add(new Paragraph("At.: Sr. Mauricio Almanza\n\n", smallerFont));
        document.Add(new Paragraph("Ref.: Solicitudes de Garantía- Sistema SAGA/2\n\n", smallerFont));
    }

    private void AddDynamicText(Document document, string filePath)
    {
        Font smallerFont = new Font(Font.FontFamily.HELVETICA, 10); // Smaller font size

        var lines = File.ReadAllLines(filePath);

        if (lines.Length > 0)
        {
            // Assuming the lines array contains the necessary data in a specific format
            string[] data = lines[0].Split(' ');

            document.Add(new Paragraph($"Lista de Credito {data[0]}\n\n", smallerFont));
            document.Add(new Paragraph($"Mano de obra ........................................................ US$ {data[1]}\n", smallerFont));
            document.Add(new Paragraph($"Repuestos (incl. L. Cost).......................................... US$ {data[2]}\n", smallerFont));
            document.Add(new Paragraph($"Total........................................................................ US$ {data[3]}\n", smallerFont));
            document.Add(new Paragraph("Esta lista corresponde a:\n", smallerFont));
            document.Add(new Paragraph($"Solicitudes ref. movimiento de {data[4]}.\n", smallerFont));
            document.Add(new Paragraph($"Débito {data[5]} = US$ {data[6]}\n\n", smallerFont));
        }

        // Add the image below the dynamic text
        string footerImagePath = ConfigurationManager.AppSettings["FooterImagePath"];
        if (File.Exists(footerImagePath))
        {
            Image footerImage = Image.GetInstance(footerImagePath);
            footerImage.ScaleToFit(420f, 420f); // Adjust the size to be larger
            footerImage.Alignment = Element.ALIGN_LEFT; // Align the image to the left
            document.Add(footerImage);
        }
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
                Subject = "New PDF Invoice",
                Body = "Please find the attached PDF invoice."
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
