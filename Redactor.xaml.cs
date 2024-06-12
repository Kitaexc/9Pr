using Microsoft.Win32;
using System.IO;
using System.Windows;
using System.Windows.Documents;
using Spire.Doc;
using Spire.Doc.Documents;

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

        public void LoadFile(string filePath)
        {
            string fileExtension = Path.GetExtension(filePath).ToLower();
            if (fileExtension == ".rtf")
            {
                LoadRtfFile(filePath);
            }
            else if (fileExtension == ".docx")
            {
                LoadDocxFile(filePath);
            }
            else
            {
                MessageBox.Show("Unsupported file format", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LoadRtfFile(string filePath)
        {
            TextRange range = new TextRange(RichTextBox.Document.ContentStart, RichTextBox.Document.ContentEnd);
            using (FileStream fileStream = new FileStream(filePath, FileMode.Open))
            {
                range.Load(fileStream, DataFormats.Rtf);
            }
        }

        private void LoadDocxFile(string filePath)
        {
            // Создание нового документа и загрузка содержимого из файла
            Document document = new Document();
            document.LoadFromFile(filePath);

            // Очистка текущего содержимого RichTextBox
            RichTextBox.Document.Blocks.Clear();

            // Перебор всех разделов документа
            foreach (Spire.Doc.Section section in document.Sections)
            {
                // Перебор всех параграфов в каждом разделе
                foreach (Spire.Doc.Documents.Paragraph paragraph in section.Paragraphs)
                {
                    // Создание нового абзаца
                    System.Windows.Documents.Paragraph newParagraph = new System.Windows.Documents.Paragraph();

                    // Перебор всех элементов в параграфе
                    foreach (DocumentObject docObject in paragraph.ChildObjects)
                    {
                        // Если элемент является текстом
                        if (docObject is Spire.Doc.Fields.TextRange textRange)
                        {
                            Run run = new Run(textRange.Text);

                            // Применение форматирования
                            run.FontWeight = textRange.CharacterFormat.Bold ? FontWeights.Bold : FontWeights.Normal;
                            run.FontStyle = textRange.CharacterFormat.Italic ? FontStyles.Italic : FontStyles.Normal;
                            run.TextDecorations = textRange.CharacterFormat.UnderlineStyle != UnderlineStyle.None ? TextDecorations.Underline : null;

                            newParagraph.Inlines.Add(run);
                        }
                    }

                    // Добавление абзаца в RichTextBox
                    RichTextBox.Document.Blocks.Add(newParagraph);
                }
            }
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
                System.Windows.Documents.TextRange range = new System.Windows.Documents.TextRange(RichTextBox.Document.ContentStart, RichTextBox.Document.ContentEnd);
                range.Save(fileStream, DataFormats.Rtf);
            }
            MessageBox.Show("Файл успешно сохранен как RTF!", "Сохранение", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void SaveAsDocx(string fileName)
        {
            // Создание нового документа
            Document document = new Document();

            // Создание нового раздела и добавление его в документ
            Spire.Doc.Section section = document.AddSection();

            // Создание нового абзаца
            Spire.Doc.Documents.Paragraph paragraph = section.AddParagraph();

            // Получение текста из RichTextBox
            System.Windows.Documents.TextRange textRange = new System.Windows.Documents.TextRange(RichTextBox.Document.ContentStart, RichTextBox.Document.ContentEnd);

            // Добавление текста в абзац
            paragraph.AppendText(textRange.Text);

            // Сохранение документа
            document.SaveToFile(fileName, FileFormat.Docx);

            MessageBox.Show("Файл успешно сохранен как DOCX!", "Сохранение", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            SendWindow sending = new SendWindow(RichTextBox);
            sending.Show();
        }
    }
}
