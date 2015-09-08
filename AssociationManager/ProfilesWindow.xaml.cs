// Copyright (c) 2015, Dijji, and released under Ms-PL.  This, with other relevant licenses, can be found in the root of this distribution.

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
using FileMetadataAssociationManager.Resources;

namespace FileMetadataAssociationManager
{
    /// <summary>
    /// Interaction logic for ProfilesWindow.xaml
    /// </summary>
    public partial class ProfilesWindow : Window
    {
        private State state;
        private ProfilesView view;

        public ProfilesWindow(State state)
        {
            InitializeComponent();
            this.state = state;

            if (state.AllProperties.Count == 0)
                state.PopulateSystemProperties();

            view = new ProfilesView(this, state);
            this.DataContext = view;
        }

        public ProfileControls IdentifyControl(UIElement control)
        {
            if (control is TreeView)
            {
                var tv = control as TreeView;
                if (tv == treeFullDetails)
                    return ProfileControls.FullDetails;
                if (tv == treeAllProperties)
                    return ProfileControls.Properties;
            }
            else if (control is ListBox)
            {
                var lb = control as ListBox;
                if (lb == listPreviewDetails)
                    return ProfileControls.PreviewDetails;
                if (lb == listPropGroup)
                    return ProfileControls.Groups;
            }

            return ProfileControls.Unknown;
        }

        private void useDragDrop_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show(LocalizedMessages.UseDragDrop, LocalizedMessages.HelpHeader, MessageBoxButton.OK);
        }

        private void cbApply_Click(object sender, RoutedEventArgs e)
        {
            view.ApplyChanges();
        }

        private void cbCancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void cbOK_Click(object sender, RoutedEventArgs e)
        {
            if (view.ApplyChanges())
                this.Close();
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            // Pass back the currently selected profile, if there is one
            state.SelectedProfile = view.SelectedProfile != null ? view.SelectedProfile.Original : null;
        }

        private void treeProfiles_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            TreeView tv = (TreeView)sender;
            TreeItem ti = (TreeItem)tv.SelectedItem;
            view.IsInTreeTextBoxEditMode = false;
            view.SelectedProfile = (Profile)ti.Item;
        }

        private void New_CanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            TreeItem ti = DetermineTreeItem(e);
            bool canExecute = true;

            if (ti != null)
            {
                Profile p = ti.Item as Profile;
                canExecute = p == null ? ti.Name == LocalizedMessages.ProfileCustom : p.IsReadOnly == false;
            }

            e.CanExecute = canExecute;
            if (!canExecute)
                e.Handled = true;
        }

        private void New_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            view.AddNewProfile();
        }

        private void Clone_CanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            TreeItem ti = DetermineTreeItem(e);
            bool canExecute = false;

            if (ti != null)
            {
                Profile p = ti.Item as Profile;
                canExecute = p != null;
            }

            e.CanExecute = canExecute;
            if (!canExecute)
                e.Handled = true;
        }

        private void Clone_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            TreeItem ti = DetermineTreeItem(e);
            Profile p = ti.Item as Profile;

            if (p != null)
            {
                view.AddNewProfile(p);
            }
        }

        private void Delete_CanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            TreeItem ti = DetermineTreeItem(e);
            bool canExecute = false;

            if (ti != null)
            {
                Profile p = ti.Item as Profile;
                canExecute = view.IsProfileWritable(p);
            }

            e.CanExecute = canExecute;
            if (!canExecute)
                e.Handled = true;

        }

        private void Delete_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            TreeItem ti = DetermineTreeItem(e);
            Profile p = ti.Item as Profile;

            if (view.IsProfileWritable(p))
                view.DeleteProfile(p);
        }

        private void Rename_CanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            TreeItem ti = DetermineTreeItem(e);
            bool canExecute = false;

            if (ti != null)
            {
                Profile p = ti.Item as Profile;
                canExecute = view.IsProfileWritable(p);
            }

            e.CanExecute = canExecute;
            if (!canExecute)
                e.Handled = true;
        }

        private void Rename_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            view.IsInTreeTextBoxEditMode = true;
        }

        private void ToggleEdit_CanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            bool canExecute = false;

            if (view.IsInTreeTextBoxEditMode)
                // Can always leave edit mode
                canExecute = true;

            else
            {
                TreeItem ti = DetermineTreeItem(e);

                if (ti != null)
                {
                    Profile p = ti.Item as Profile;
                    canExecute = view.IsProfileWritable(p);
                }
             }

            e.CanExecute = canExecute;
            if (!canExecute)
                e.Handled = true;
        }

        private void ToggleEdit_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            view.IsInTreeTextBoxEditMode = !view.IsInTreeTextBoxEditMode;
        }

        private void RemoveFullDetails_CanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            TreeItem ti = DetermineTreeItem(e);
            bool canExecute = view.IsSelectedProfileWritable && ti != null;

            e.CanExecute = canExecute;
            if (!canExecute)
                e.Handled = true;
        }

        private void RemoveFullDetails_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            TreeItem ti = DetermineTreeItem(e);
            view.RemoveFullDetailsItem(ti);
        }

        private void RemovePreviewDetails_CanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            string property = DetermineProperty(e);
            bool canExecute = view.IsSelectedProfileWritable && property != null;

            e.CanExecute = canExecute;
            if (!canExecute)
                e.Handled = true;
        }

        private void RemovePreviewDetails_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            string property = DetermineProperty(e);
            view.RemovePreviewDetailsItem(property);
        }

        private TreeItem DetermineTreeItem(CanExecuteRoutedEventArgs e)
        {
            TreeItem ti = null;

            if (e.Source is TreeView)
            {
                if (e.OriginalSource is TreeViewItem)
                {
                    // Take the item on which the context menu was invoked
                    TreeViewItem tvi = (TreeViewItem)e.OriginalSource;
                    ti = (TreeItem)tvi.Header;
                }
                else
                {
                    // TreeView
                    // A shortcut key was used, so take the currently selected item
                    if (e.Parameter as string != null && ((string)e.Parameter) == "Keystroke")
                    {
                        ti = (TreeItem)((TreeView)e.Source).SelectedItem;
                    }
                }
            }

            return ti;
        }

        private TreeItem DetermineTreeItem(ExecutedRoutedEventArgs e)
        {
            if (e.OriginalSource is TreeViewItem)
            {
                TreeViewItem tvi = (TreeViewItem)e.OriginalSource;
                return (TreeItem)tvi.Header;
            }
            else
            {
                // TreeView
                return (TreeItem)((TreeView)e.Source).SelectedItem;
            }
        }

        private string DetermineProperty(CanExecuteRoutedEventArgs e)
        {
            string property = null;

            if (e.Source is ListBox)
            {
                // Take the item on which the context menu was invoked
                if (e.OriginalSource is ListBoxItem)
                {
                    ListBoxItem lbi = (ListBoxItem)e.OriginalSource;
                    property = (string)lbi.Content;
                }
                else
                {
                    // ListBox
                    // A shortcut key was used, so take the currently selected item
                    if (e.Parameter as string != null && ((string)e.Parameter) == "Keystroke")
                    {
                        property = (string)((ListBox)e.Source).SelectedItem;
                    }
                }
            }

            return property;
        }

        private string DetermineProperty(ExecutedRoutedEventArgs e)
        {
            if (e.OriginalSource is ListBoxItem)
            {
                ListBoxItem lbi = (ListBoxItem)e.OriginalSource;
                return (string)lbi.Content;
            }
            else
            {
                // ListBox
                return (string)((ListBox)e.Source).SelectedItem;
            }
        }

        private void TextBox_IsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            // Invoked when editable text is made visible
            // Select all of the current text
            TextBox tb = (TextBox)sender;
            tb.Focus();
            tb.SelectionStart = 0;
            tb.SelectionLength = tb.Text.Length;
        }
    }
}
