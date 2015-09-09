using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace FileMetadataAssociationManager
{
    /// <summary>
    /// Interaction logic for HowTo.xaml
    /// </summary>
    public partial class HowTo : Window
    {
        public HowTo()
        {
            InitializeComponent();
        }

        private void close_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void CheckBox_Checked(object sender, RoutedEventArgs e)
        {
            Properties.Settings.Default.SuppressHowTo = true;
            Properties.Settings.Default.Save();
        }
    }
}
