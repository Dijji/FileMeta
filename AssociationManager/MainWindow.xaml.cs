// Copyright (c) 2013, Dijii, and released under Ms-PL.  This, with other relevant licenses, can be found in the root of this distribution.

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.ComponentModel;
using FileMetadataAssociationManager.Resources;


namespace FileMetadataAssociationManager
{
     /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private State state = new State();
        
        public MainWindow()
        {
            InitializeComponent();
            this.DataContext = state;
            state.Populate();
            state.PropertyChanged += new PropertyChangedEventHandler(state_PropertyChanged);
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            if (listExtensions.Items.Count > 0)
                listExtensions.SelectedItem = listExtensions.Items[0];

            if (comboProfile.Items.Count > 0)
                comboProfile.SelectedItem = comboProfile.Items[0];
        }

        private void listExtensions_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ListBox lb = (ListBox)sender;
            state.SelectedExtension = (Extension)lb.SelectedItem;
            state.SelectedExtension.SelectCurrentProfileIfKnown();
        }

        private void comboProfile_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ComboBox cb = (ComboBox)sender;
            state.SelectedProfile = (Profile)cb.SelectedItem;
        }

        private void state_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            // Ensure that the profile combo box reflects the current selection
            if (e.PropertyName == "SelectedProfile")
            {
                if (state.SelectedProfile != comboProfile.SelectedItem)
                    comboProfile.SelectedItem = state.SelectedProfile;
            }
        }

        private void addHandler_Click(object sender, RoutedEventArgs e)
        {
            bool success = false;

            if (state.SelectedExtension != null && state.SelectedProfile != null)
                 success = state.SelectedExtension.SetupHandlerForExtension(state.SelectedProfile);

            if (!success)
                MessageBox.Show(LocalizedMessages.HandlerSetupIssues);
        }

        private void removeHandler_Click(object sender, RoutedEventArgs e)
        {
            if (state.SelectedExtension != null)
                state.SelectedExtension.RemoveHandlerFromExtension();
        }
    }
}


