using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;

namespace Word
{
    public partial class SendWindow : Window
    {
        private RichTextBox _richTextBox;

        public SendWindow(RichTextBox richTextBox)
        {
            InitializeComponent();
            this.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            this.MinWidth = 450;
            this.MinHeight = 270;

            _richTextBox = richTextBox;
        }

        private void SendButton_Click(object sender, RoutedEventArgs e)
        {
            string tempFilePath = SaveRichTextBoxContentToTempFile();
            if (tempFilePath != null)
            {
                SendEmailWithAttachment(Login.Text, Password.Password, Login_Friend.Text, Topic.Text, tempFilePath);
            }
        }

        private string SaveRichTextBoxContentToTempFile()
        {
            try
            {
                string tempFileName = Path.GetTempFileName();
                string tempFilePath = Path.ChangeExtension(tempFileName, ".docx"); // или ".rtf" для RTF формата

                SaveAsDocx(tempFilePath);
                return tempFilePath;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при сохранении файла: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }
        }

        private void SaveAsDocx(string fileName)
        {
            // Создание нового документа
            Spire.Doc.Document document = new Spire.Doc.Document();

            // Создание нового раздела и добавление его в документ
            Spire.Doc.Section section = document.AddSection();

            // Создание нового абзаца
            Spire.Doc.Documents.Paragraph paragraph = section.AddParagraph();

            // Получение текста из RichTextBox
            System.Windows.Documents.TextRange textRange = new System.Windows.Documents.TextRange(_richTextBox.Document.ContentStart, _richTextBox.Document.ContentEnd);

            // Добавление текста в абзац
            paragraph.AppendText(textRange.Text);

            // Сохранение документа
            document.SaveToFile(fileName, Spire.Doc.FileFormat.Docx);
        }

        private void SendEmailWithAttachment(string fromEmail, string password, string toEmail, string subject, string attachmentFilePath)
        {
            try
            {
                SmtpClient client = GetSmtpClient(fromEmail, password);

                MailMessage mailMessage = new MailMessage
                {
                    From = new MailAddress(fromEmail),
                    Subject = subject,
                    Body = "Вложение содержит ваш файл.",
                    IsBodyHtml = true,
                };

                mailMessage.To.Add(toEmail);
                mailMessage.Attachments.Add(new Attachment(attachmentFilePath));

                client.Send(mailMessage);

                MessageBox.Show("Письмо успешно отправлено!", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при отправке письма: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private SmtpClient GetSmtpClient(string fromEmail, string password)
        {
            SmtpClient client = new SmtpClient();
            string domain = fromEmail.Split('@')[1];

            switch (domain.ToLower())
            {
                case "gmail.com":
                    client.Host = "smtp.gmail.com";
                    client.Port = 587;
                    client.EnableSsl = true;
                    break;

                case "mail.ru":
                    client.Host = "smtp.mail.ru";
                    client.Port = 587;
                    client.EnableSsl = true;
                    break;

                case "rambler.ru":
                    client.Host = "smtp.rambler.ru";
                    client.Port = 465;
                    client.EnableSsl = true;
                    break;

                case "yandex.ru":
                case "yandex.com":
                    client.Host = "smtp.yandex.ru";
                    client.Port = 465;
                    client.EnableSsl = true;
                    break;

                default:
                    throw new NotSupportedException($"Домен {domain} не поддерживается для отправки почты.");
            }

            client.Credentials = new NetworkCredential(fromEmail, password);
            return client;
        }
    }
}
