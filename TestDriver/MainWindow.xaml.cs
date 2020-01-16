// Copyright (c) 2013, Dijji, and released under Ms-PL.  This, with other relevant licenses, can be found in the root of this distribution.
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.ComponentModel;
using TestDriver.Resources;

// DSOFile registration (e.g. using regsvr32) is a pre-requisite for some of these tests

// VS 2010 requires a workaround to import a 32-bit DLL like DSOFile into a 64-bit project.  The following was used to create the Interop DLL in \lib
// C:\Program Files\Microsoft SDKs\Windows\v7.1\Bin\x64>tlbimp "C:\Test\FileMetadata\Test Driver\lib\dsofile.dll" /namespace:DSOFile /machine:x86 /out:"C:\Test\FileMetadata\Test Driver\lib\Interop.DSOFile.dll"


namespace TestDriver
{
     /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private State state = null;
        
        public MainWindow()
        {
            InitializeComponent();
            state = new State(this); 
            this.DataContext = state;
            state.Populate();
            state.PropertyChanged += new PropertyChangedEventHandler(state_PropertyChanged);
        }

        private void RunAll_Click(object sender, RoutedEventArgs e)
        {
            if (state.Running == RunState.Idle)
            {
                state.Status = "Running all tests...";
                state.Running = RunState.Running;

                state.Results.Clear();
                state.TestsToRun.Clear();
                foreach (Test t in lbTests.Items)
                {
                    state.TestsToRun.Add(t);
                }

                ThreadPool.QueueUserWorkItem(RunTests, state);
            }
        }

        private void RunSelected_Click(object sender, RoutedEventArgs e)
        {
            if (state.Running == RunState.Idle)
            {
                state.Status = "Running selected tests...";
                state.Running = RunState.Running;

                state.Results.Clear();
                state.TestsToRun.Clear();

                foreach (Test t in lbTests.SelectedItems)
                {
                    state.TestsToRun.Add(t);
                }

                ThreadPool.QueueUserWorkItem(RunTests, state);
            }
        }

        private void OneOff_Click(object sender, RoutedEventArgs e)
        {
            if (state.Running == RunState.Idle)
            {
                state.Status = "Running One-off tests...";
                state.Running = RunState.Running;

                state.Results.Clear();

                // Wire this up to whatever it is we think we need to run
                
                ThreadPool.QueueUserWorkItem(RunPropertyTests, state);
            }
        }

        private void LoopAll_Click(object sender, RoutedEventArgs e)
        {
            if (state.Running == RunState.Idle)
            {
                state.Status = "Looping...";
                state.Running = RunState.Looping;
                LoopAll.Content = "Stop Looping";

                state.Results.Clear();
                state.TestsToRun.Clear();
                foreach (Test t in lbTests.Items)
                {
                    state.TestsToRun.Add(t);
                }

                ThreadPool.QueueUserWorkItem(RunTests, state);
            }
            else if (state.Running == RunState.Looping)
            {
                state.Running = RunState.StopPending;
            }
        }

        void state_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "Running")
            {
                if (state.Running == RunState.Idle)
                {
                    state.Status = "Ready";
                    Dispatcher.Invoke(new Action(() => LoopAll.Content = "Loop All Tests"));
                }
            }
        }

        private string resultsDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        private void SaveResults_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new System.Windows.Forms.SaveFileDialog();

            dialog.Title = "Specify file to save to";
            dialog.InitialDirectory = resultsDirectory;
            dialog.Filter = "txt Files (*.txt)|*.txt";
            dialog.FileName = "results";

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                resultsDirectory = Path.GetDirectoryName(dialog.FileName);
                using(var tw = new StreamWriter(dialog.FileName))
                {
                    foreach (var r in state.Results)
                        tw.WriteLine(r);
                }
            }
        }

        private static void RunPropertyTests(Object state)
        {
            State s = (State)state;
            List<string> summary = new List<string>();

            // Iterate over the files in a hard coded folder,
            // pumping as many properties as possible into each of them, and recording the results
            string targetDir = @"D:\TestProperties";
            string resultsDir = Path.Combine(targetDir, "Results");
            Directory.CreateDirectory(resultsDir);  // Ensure all directories exist
            foreach (string fileName in Directory.GetFiles(targetDir))
            {
                var ext = Path.GetExtension(fileName);
                var test = new TestMassRoundTrip(ext, fileName);
                test.Run(s);
                using (var tw = new StreamWriter(resultsDir + @"\Properties" + ext + ".txt"))
                {
                    foreach (var r in s.Results)
                        tw.WriteLine(r);
                }
                summary.Add(String.Format("{0}, {1}", ext, s.Results[s.Results.Count() - 2]));
                s.window.Dispatcher.Invoke(new Action(() => s.Results.Clear()));
            }
            using (var tw = new StreamWriter(resultsDir + @"\Properties Summary.csv"))
            {
                foreach (var r in summary)
                    tw.WriteLine(r);
            } 
            s.window.Dispatcher.Invoke(new Action(() => { s.Results.Clear(); s.Status = "Done"; }));
            s.Running = RunState.Idle;
        }

        private static void RunTests(Object state)
        {
            State s = (State)state;
            int loops = 0;

            while (true)
            {
                foreach (Test t in s.TestsToRun)
                {
                    if (s.Running == RunState.StopPending)
                        break;

                    if (!t.Run(s))
                        s.window.Dispatcher.Invoke(new Action(() => s.Running = RunState.StopPending));
                }

                if (s.Running != RunState.Looping)
                    break;
                else
                    s.window.Dispatcher.Invoke(new Action(() => { s.Results.Clear(); s.Status = (++loops).ToString() + " loops completed..."; }));
            }
            s.Running = RunState.Idle;
        }

        private void CopyListBoxItem_Click(object sender, RoutedEventArgs e)
        {
            if (lbLog.SelectedItem != null)
                Clipboard.SetText(lbLog.SelectedItem.ToString());
        }

        private void CommandBinding_CanExecute(object sender, System.Windows.Input.CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = lbLog.HasItems;
        }

        private void CommandBinding_Executed(object sender, System.Windows.Input.ExecutedRoutedEventArgs e)
        {
        }
     }
}


