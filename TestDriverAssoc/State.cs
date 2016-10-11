// Copyright (c) 2016, Dijji, and released under Ms-PL.  This, with other relevant licenses, can be found in the root of this distribution.

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Automation;

namespace TestDriverAssoc
{
    class State
    {
        private ObservableCollection<string> reports = new ObservableCollection<string>();

        public FileInfo BackupFileInfo { get; set; }
        public bool WasNoState { get; set; }
        public int RunCounter { get; set; }
        public int TestCounter { get; set; } 
        public ObservableCollection<string> Reports { get { return reports; } }

        // Used for GUI testing
        public Process process;
        public AutomationElement mainWindow = null;
        public AutomationElement extensions = null;
        public AutomationElement profiles = null;
        public AutomationElement addButton = null;
        public AutomationElement removeButton = null;
        public AutomationElement updateButton = null;

        public void AddReport(string text)
        {
            Application.Current.Dispatcher.Invoke(new Action(delegate
            {
                Reports.Add(text);
            }));
        }

    }
}
