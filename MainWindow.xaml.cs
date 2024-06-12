using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace Word
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            this.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            this.MinWidth = 250;
            this.MinHeight = 400;
        }

        private void Open_Word_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Word Document (*.docx)|*.docx|Rich Text Format (*.rtf)|*.rtf|All files (*.*)|*.*";
            if (openFileDialog.ShowDialog() == true)
            {
                Redactor redactor = new Redactor();
                redactor.LoadFile(openFileDialog.FileName);
                redactor.Show();
                this.Close();
            }
        }


        private void Create_Word_Click(object sender, RoutedEventArgs e)
        {
            Redactor redactor = new Redactor();
            redactor.Show();
            this.Close();
        }

    }
}
