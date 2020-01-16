// Copyright (c) 2013, Dijji, and released under Ms-PL.  This, with other relevant licenses, can be found in the root of this distribution.

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.ComponentModel;
using AssociationMessages;

namespace FileMetadataAssociationManager
{
     /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private State state = new State();
        private MainView view;
        
        public MainWindow()
        {
            InitializeComponent();

            view = new MainView(this, state);
            this.DataContext = view;
            view.PropertyChanged += new PropertyChangedEventHandler(view_PropertyChanged);

            try
            {
                state.Populate();
            }
            catch (AssocMgrException ae)
            {
                System.Windows.MessageBox.Show(ae.DisplayString, LocalizedMessages.ErrorHeader);
            }
         }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            if (listExtensions.Items.Count > 0)
                listExtensions.SelectedItem = listExtensions.Items[0];

            if (comboProfile.Items.Count > 0 && comboProfile.SelectedItem == null)
                comboProfile.SelectedItem = comboProfile.Items[0];
        }

        private void Window_Closing(object sender, CancelEventArgs e)
        {
            // If there have been changes to registry settings, offer to restart Explorer
            if (state.HasChanged)
            {
                var result = MessageBox.Show(LocalizedMessages.RestartExplorerNow, LocalizedMessages.ClosingWithChanges, MessageBoxButton.YesNoCancel);

                if (result == MessageBoxResult.No)
                    return;
                else if (result == MessageBoxResult.Cancel)
                {
                    e.Cancel = true;
                    return;
                }
                else
                {
                    restartExplorer_Click(null, null);
                }
            }
        }

        private void listExtensions_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ListView lv = (ListView)sender;
            view.SetSelectedExtensions(lv.SelectedItems);
         }

        private void comboProfile_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ComboBox cb = (ComboBox)sender;
            view.SelectedProfile = (Profile)cb.SelectedItem;
        }

        private void view_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            // Ensure that the profile combo box reflects the current selection
            if (e.PropertyName == "SelectedProfile")
            {
                if (view.SelectedProfile != comboProfile.SelectedItem)
                    comboProfile.SelectedItem = view.SelectedProfile;
            }
        }

        private void addHandler_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                view.AddHandlers();
            }
            catch (AssocMgrException ae)
            {
                System.Windows.MessageBox.Show(ae.DisplayString, LocalizedMessages.ErrorHeader);
            }
            catch (Exception ex)
            {
                Clipboard.SetText(ex.ToString());
                System.Windows.MessageBox.Show(ex.ToString(), LocalizedMessages.ErrorHeader);
            }
        }

        private void removeHandler_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                view.RemoveHandlers();
            }
            catch (AssocMgrException ae)
            {
                System.Windows.MessageBox.Show(ae.DisplayString, LocalizedMessages.ErrorHeader);
            }
            catch (Exception ex)
            {
                Clipboard.SetText(ex.ToString());
                System.Windows.MessageBox.Show(ex.ToString(), LocalizedMessages.ErrorHeader);
            }
        }

        private void refresh_Click(object sender, RoutedEventArgs e)
        {
            // Re-sort the handlers, while maintaining the current selections
            List<Extension> currentSelections = new List<Extension>(listExtensions.SelectedItems.Cast<Extension>());
            state.SortExtensions();
            foreach (var sel in currentSelections)
                listExtensions.SelectedItems.Add(sel);
            listExtensions.ScrollIntoView(currentSelections.FirstOrDefault());
            view.SortRequired = false;
        }

        private void profiles_Click(object sender, RoutedEventArgs e)
        {
            // Open the custom profiles dialog
            state.SelectedProfile = view.SelectedProfile;
            var w = new ProfilesWindow(state);
            w.Owner = this;
            w.ShowDialog();
            view.RefreshProfiles();
            view.SelectedProfile = state.SelectedProfile;
        }

        // Temporary code to enumerate extensions for which Add Handler fails
        //private void CheckForBlockedExtensions()
        //{
        //    using (var w = new System.IO.StreamWriter(@"D:\TestProperties\Results\Extensions.csv"))
        //    {
        //        foreach (var ext in state.Extensions.Where(x => x.PropertyHandlerState == HandlerState.Foreign))
        //        {
        //            string result = "Exception";
        //            try
        //            {
        //                result = ext.IsPropertyHandlerBlocked() ? "Blocked" : "Ok";
        //            }
        //            catch (System.Exception ex)
        //            {
        //            }
        //            w.WriteLine(String.Format("{0}, {1}, {2}", ext.Name, result, ext.PropertyHandlerDisplay));
        //        }
        //    }
        //}

        private void restartExplorer_Click(object sender, RoutedEventArgs e)
        {
            // Temporary code to enumerate extensions for which Add Handler fails
            //CheckForBlockedExtensions();
            //return;

            bool failed = false;
            try
            {
                foreach (Process p in Process.GetProcessesByName("explorer"))
                {
                    p.Kill();
                    p.WaitForExit(); // possibly with a timeout
                }
            }
            catch (System.Exception ex)
            {
                failed = true;
            }

            // If that worked, start up a new instance so that at least the desktop is available
            if (!failed)
            {
                Process p = new Process();
                p.StartInfo.UseShellExecute = false;
                p.StartInfo.FileName = Environment.GetEnvironmentVariable("SystemRoot") + @"\explorer.exe";
                p.Start();
                state.HasChanged = false;
            }
        }

        private void updateRegistry_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                view.RefreshRegistry();
            }
            catch (AssocMgrException ae)
            {
                System.Windows.MessageBox.Show(ae.DisplayString, LocalizedMessages.ErrorHeader);
            }
        }
    }
}


