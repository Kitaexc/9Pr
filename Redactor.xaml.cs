using Microsoft.Win32;
using System.IO;
using System.Windows;
using System.Windows.Documents;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlParagraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using OpenXmlRun = DocumentFormat.OpenXml.Wordprocessing.Run;
using OpenXmlBody = DocumentFormat.OpenXml.Wordprocessing.Body;
using OpenXmlText = DocumentFormat.OpenXml.Wordprocessing.Text;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Windows.Controls;

namespace Word
{
    public partial class Redactor : Window
    {
        public Redactor()
        {
            InitializeComponent();
            this.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            this.MinWidth = 560;
            this.MinHeight = 600;
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Rich Text Format (*.rtf)|*.rtf|Word Document (*.docx)|*.docx|All files (*.*)|*.*";
            if (saveFileDialog.ShowDialog() == true)
            {
                string fileExtension = Path.GetExtension(saveFileDialog.FileName).ToLower();
                if (fileExtension == ".rtf")
                {
                    SaveAsRtf(saveFileDialog.FileName);
                }
                else if (fileExtension == ".docx")
                {
                    SaveAsDocx(saveFileDialog.FileName);
                }
                else
                {
                    MessageBox.Show("Unsupported file format", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void SaveAsRtf(string fileName)
        {
            using (FileStream fileStream = new FileStream(fileName, FileMode.Create))
            {
                TextRange range = new TextRange(RichTextBox.Document.ContentStart, RichTextBox.Document.ContentEnd);
                range.Save(fileStream, DataFormats.Rtf);
            }
            MessageBox.Show("Файл успешно сохранен как RTF!", "Сохранение", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void SaveAsDocx(string fileName)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(fileName, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
                mainPart.Document = new Document();
                OpenXmlBody body = new OpenXmlBody();

                TextRange textRange = new TextRange(RichTextBox.Document.ContentStart, RichTextBox.Document.ContentEnd);
                OpenXmlParagraph para = new OpenXmlParagraph();
                OpenXmlRun run = new OpenXmlRun();
                run.Append(new OpenXmlText(textRange.Text));
                para.Append(run);
                body.Append(para);

                mainPart.Document.Append(body);
                mainPart.Document.Save();
            }
            MessageBox.Show("Файл успешно сохранен как DOCX!", "Сохранение", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            SendWindow sending = new SendWindow();
            sending.Show();
        }
    }
}
