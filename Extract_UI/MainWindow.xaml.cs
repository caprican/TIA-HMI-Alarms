using System;
using System.ComponentModel;
using System.Windows;

namespace Extract_UI
{
    /// <summary>
    /// Logique d'interaction pour MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        

        public bool ClosedByUser { get; set; }

        public MainWindow()
        {
            InitializeComponent();
            ClosedByUser = false;

            Closed += MainWindow_Closed;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void MainWindow_Closed(object sender, EventArgs e)
        {
            ClosedByUser = true;
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            this.Close();
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }


}
