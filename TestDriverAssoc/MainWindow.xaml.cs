// Copyright (c) 2016, Dijji, and released under Ms-PL.  This, with other relevant licenses, can be found in the root of this distribution.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Win32;

namespace TestDriverAssoc
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private State state = new State();
        private Dictionary<string, RegState> regStates = new Dictionary<string, RegState>();


        public MainWindow()
        {
            InitializeComponent();
            DataContext = state;
            state.RunCounter = 1;
        }

        private void scan_Click(object sender, RoutedEventArgs e)
        {

            // get all the current handlers
            using (RegistryKey handlers = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\PropertySystem\PropertyHandlers", false))
            {
                foreach (string name in handlers.GetSubKeyNames())
                {
                    using (RegistryKey key = handlers.OpenSubKey(name, false))
                    {
                        string handlerGuid = (string)key.GetValue(null);
                        if (handlerGuid == Const.OurPropertyHandlerGuid)
                        {
                            RegState s = new RegState();
                            s.Read(name);
                            regStates.Add(name, s);
                        }
                    }
                }
            }
        }

        private void testCommandLine_Click(object sender, RoutedEventArgs e)
        {
            ThreadPool.QueueUserWorkItem(TestCommandLine.Run, state);
        }

        private void testGUI_Click(object sender, RoutedEventArgs e)
        {
            ThreadPool.QueueUserWorkItem(TestGUI.Run, state);
        }

        private void clear_Click(object sender, RoutedEventArgs e)
        {
            state.Reports.Clear();
        }

        private void save_Click(object sender, RoutedEventArgs e)
        {
            // Save the results of the last test
            System.Windows.Forms.SaveFileDialog dialog = new System.Windows.Forms.SaveFileDialog();

            dialog.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            dialog.FilterIndex = 0;
            dialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                StreamWriter sw = new StreamWriter(dialog.OpenFile());
                foreach (var r in state.Reports)
                {
                    sw.WriteLine(r);
                }
                sw.Close();
            }
        }
    }
}
