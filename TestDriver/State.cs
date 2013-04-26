// Copyright (c) 2013, Dijii, and released under the Common Public License.  This, with other relevant licenses, can be found in the root of this distribution.

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Threading;
using Microsoft.Win32;
using System.Runtime.InteropServices;
using TestDriver.Resources;
using TestDriverCodePack;
//using Microsoft.WindowsAPICodePack.Shell;
//using Microsoft.WindowsAPICodePack.Shell.PropertySystem;

namespace TestDriver
{
    public enum RunState
    {
        Idle,
        Running,
        Looping,
        StopPending
    }

    public class State : INotifyPropertyChanged
    {
        private Test selectedTest = null;
        private ObservableCollection<Test> tests = new ObservableCollection<Test>();
        private ObservableCollection<string> results = new ObservableCollection<string>();
        private string status;
        private RunState running;
        private List<ShellPropertyDescription> propDescs = new List<ShellPropertyDescription>();
        private List<Test> testsToRun = new List<Test>();
        public MainWindow window;

        public ObservableCollection<Test> Tests { get { return tests; } }
        public ObservableCollection<string> Results { get { return results; } }
        public string Status { get { return status; } set { status = value; OnPropertyChanged("Status"); } }
        public Test SelectedTest { get { return selectedTest; } set { selectedTest = value; OnPropertyChanged("SelectedTest"); } }
        public List<ShellPropertyDescription> PropertyDescriptions { get { return propDescs; } }
        public RunState Running { get { return running; } set { running = value; OnPropertyChanged("Running"); OnPropertyChanged("Idle"); OnPropertyChanged("IdleOrLooping"); } }
        public bool Idle { get { return running == RunState.Idle; } }
        public bool IdleOrLooping { get { return running == RunState.Idle || running == RunState.Looping; } }
        public List<Test> TestsToRun { get { return testsToRun; } }

        public State(MainWindow window)
        {
            this.window = window;
        }

        public void Populate()
        {
            PopulateTests();
            PopulateSystemProperties();
        }

        public void RecordEntry(string entry)
        {
            window.Dispatcher.Invoke(new Action<ObservableCollection<string>>(
                r => r.Add(entry)), Results);
        }

        public void RecordResult(string name, bool good)
        {
            window.Dispatcher.Invoke(new Action< ObservableCollection<string>>(
                r => r.Add((good ? "Pass: " : "Fail :") + name)), Results);
        }

        private void PopulateTests()
        {
            Tests.Add(new RoundTrip1());
            Tests.Add(new RoundTrip2());
            Tests.Add(new RoundTrip3());
            Tests.Add(new ExportImport1());
            Tests.Add(new TestProperties1());
            SelectedTest = Tests.Count > 0 ? Tests[0] : null;
        }

        private void PopulateSystemProperties()
        {
            IPropertyDescriptionList propertyDescriptionList = null;
            IPropertyDescription propertyDescription = null;
            Guid guid = new Guid(ShellIIDGuid.IPropertyDescriptionList);

            try
            {
                int hr = PropertySystemNativeMethods.PSEnumeratePropertyDescriptions(PropertySystemNativeMethods.PropDescEnumFilter.PDEF_ALL, ref guid, out propertyDescriptionList);
                if (hr >= 0)
                {
                    uint count;
                    propertyDescriptionList.GetCount(out count);
                    guid = new Guid(ShellIIDGuid.IPropertyDescription);
                    Dictionary<string, List<string>> dict = new Dictionary<string, List<string>>();

                    for (uint i = 0; i < count; i++)
                    {
                        propertyDescriptionList.GetAt(i, ref guid, out propertyDescription);

                        propDescs.Add(new ShellPropertyDescription(propertyDescription));
                    }
                }
            }
            finally
            {
                if (propertyDescriptionList != null)
                {
                    Marshal.ReleaseComObject(propertyDescriptionList);
                }
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged(String info)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(info));
            }
        }
    }
}
