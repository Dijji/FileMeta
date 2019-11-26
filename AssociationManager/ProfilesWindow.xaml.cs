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
using AssociationMessages;
using SearchTextBox;

namespace FileMetadataAssociationManager
{
    /// <summary>
    /// Interaction logic for ProfilesWindow.xaml
    /// </summary>
    public partial class ProfilesWindow : Window
    {
        private State state;
        private ProfilesView view;
        private TreeItem[] tiArray = null;
        private int searchIndex = -1;

        public ProfilesWindow(State state)
        {
            InitializeComponent();
            this.state = state;

            if (state.AllProperties.Count == 0)
                state.PopulateSystemProperties();

            view = new ProfilesView(this, state);
            this.DataContext = view;

            // Supply the Search control with the list of sections
            searchTextBox.SectionsList = new List<string> { "System", "Display" }; ;
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
                if (lb == listInfoTip)
                    return ProfileControls.InfoTip;
                if (lb == listPropGroup)
                    return ProfileControls.Groups;
            }

            return ProfileControls.Unknown;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            if (!Properties.Settings.Default.SuppressHowTo)
            {
                var w = new HowTo();
                w.Owner = this;
                w.Show();
            }
        }

        private void cbApply_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                view.ApplyChanges();
            }
            catch (AssocMgrException ae)
            {
                System.Windows.MessageBox.Show(ae.DisplayString, LocalizedMessages.ErrorHeader);
            }
        }

        private void cbCancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void cbOK_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (view.ApplyChanges())
                    this.Close();
            }
            catch (AssocMgrException ae)
            {
                System.Windows.MessageBox.Show(ae.DisplayString, LocalizedMessages.ErrorHeader);
            }
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            // Pass back the currently selected profile, if there is one
            state.SelectedProfile = view.SelectedProfile != null ? view.SelectedProfile.Original : null;
        }

        void TreeViewItem_PreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            TreeViewItem treeViewItem =
                      VisualUpwardSearch<TreeViewItem>(e.OriginalSource as DependencyObject);

            if (treeViewItem != null)
            {
                //System.Diagnostics.Trace.WriteLine(String.Format("preview selecting {0}", ((TreeItem)treeViewItem.Header).Name));
                treeViewItem.Focus();
                treeViewItem.IsSelected = true;
                e.Handled = true;
            }
        }

        static T VisualUpwardSearch<T>(DependencyObject source) where T : DependencyObject
        {
            DependencyObject returnVal = source;

            while (returnVal != null && !(returnVal is T))
            {
                DependencyObject tempReturnVal = null;
                if (returnVal is Visual) // || returnVal is Visual3D)
                {
                    tempReturnVal = VisualTreeHelper.GetParent(returnVal);
                }
                if (tempReturnVal == null)
                {
                    returnVal = LogicalTreeHelper.GetParent(returnVal);
                }
                else returnVal = tempReturnVal;
            }

            return returnVal as T;
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

        private void ToggleAsteriskFullDetails_CanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            TreeItem ti = DetermineTreeItem(e);
            bool canExecute = view.IsSelectedProfileWritable && ti != null && (PropType)ti.Item == PropType.Normal;

            e.CanExecute = canExecute;
            if (!canExecute)
                e.Handled = true;
        }

        private void ToggleAsteriskFullDetails_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            TreeItem ti = DetermineTreeItem(e);
            view.ToggleAsteriskFullDetailsItem(ti);
        }

        private void PropertyChange_CanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            PropertyListEntry property = DetermineProperty(e);
            bool canExecute = view.IsSelectedProfileWritable && property != null;

            e.CanExecute = canExecute;
            if (!canExecute)
                e.Handled = true;
        }

        private void RemovePreviewDetails_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            PropertyListEntry property = DetermineProperty(e);
            view.RemovePreviewDetailsItem(property);
        }

        private void RemoveInfoTip_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            PropertyListEntry property = DetermineProperty(e);
            view.RemoveInfoTipItem(property);
        }

        private void ToggleAsterisk_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            PropertyListEntry property = DetermineProperty(e);
            view.ToggleAsterisk(property);
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
                    //System.Diagnostics.Trace.WriteLine(String.Format("Tree at {0}", ((TreeItem)tvi.Header).Name));
                    ti = (TreeItem)tvi.Header;
                }
                else
                {
                    // TreeView
                    // If a shortcut key was used, take the currently selected item
                    if (e.Parameter as string != null && ((string)e.Parameter) == "Keystroke")
                    {
                        //System.Diagnostics.Trace.WriteLine(String.Format("Tree selected {0}", ((TreeItem)((TreeView)e.Source).SelectedItem).Name));
                        ti = (TreeItem)((TreeView)e.Source).SelectedItem;
                    }
                    //else
                    //    System.Diagnostics.Trace.WriteLine(String.Format("Tree none {0}", ""));

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

        private PropertyListEntry DetermineProperty(CanExecuteRoutedEventArgs e)
        {
            PropertyListEntry property = null;

            if (e.Source is ListBox)
            {
                // Take the item on which the context menu was invoked
                if (e.OriginalSource is ListBoxItem)
                {
                    ListBoxItem lbi = (ListBoxItem)e.OriginalSource;
                    property = (PropertyListEntry)lbi.Content;
                }
                else
                {
                    // ListBox
                    // A shortcut key was used, so take the currently selected item
                    if (e.Parameter as string != null && ((string)e.Parameter) == "Keystroke")
                    {
                        property = (PropertyListEntry)((ListBox)e.Source).SelectedItem;
                    }
                }
            }

            return property;
        }

        private PropertyListEntry DetermineProperty(ExecutedRoutedEventArgs e)
        {
            if (e.OriginalSource is ListBoxItem)
            {
                ListBoxItem lbi = (ListBoxItem)e.OriginalSource;
                return (PropertyListEntry)lbi.Content;
            }
            else
            {
                // ListBox
                return (PropertyListEntry)((ListBox)e.Source).SelectedItem;
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

        private void searchTextBox_OnSearch(object sender, RoutedEventArgs e)
        {
            try
            {
                // Obtain the property array if we have not already done so
                if (tiArray == null)
                {
                    var tiList = new List<TreeItem>();

                    foreach (var ti in state.AllProperties)
                        BuildTreeItemList(tiList, ti);

                    tiArray = tiList.ToArray();
                }

                var args = e as SearchEventArgs;
                bool system = args.Sections.Contains("System");
                bool display = args.Sections.Contains("Display");
                bool found = false;
                switch (args.SearchEventType)
                {
                    case SearchEventType.Search:
                        for (int i = 0; i < tiArray.Count(); i++)
                        {
                            found = PropertyHitTest(i, args.Keyword, system, display);
                            if (found)
                                break;
                        }
                        if (!found)
                            searchIndex = -1;
                        break;
                    case SearchEventType.Next:
                        for (int i = searchIndex + 1; i < tiArray.Count(); i++)
                        {
                            found = PropertyHitTest(i, args.Keyword, system, display);
                            if (found)
                            {
                                searchTextBox.ShowSearch = true;
                                break;
                            }
                        }
                        break;
                    case SearchEventType.Previous:
                        for (int i = searchIndex - 1; i >= 0; i--)
                        {
                            found = PropertyHitTest(i, args.Keyword, system, display);
                            if (found)
                            {
                                searchTextBox.ShowSearch = true;
                                break;
                            }
                        }
                        break;
                }

                if (!found)
                    searchTextBox.IndicateSearchFailed(args.SearchEventType);
            }
            catch (Exception ex)
            {
                // Unclear what we can do here, as we were invoked by an event from the search text box control
            }
        }

        private void treeAllProperties_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            TreeView tv = (TreeView)sender;
            TreeItem ti = (TreeItem)tv.SelectedItem; 

            if (tiArray != null)
            {
                int index = 0;
                for (; index < tiArray.Count(); index++)
                {
                    if (tiArray[index] == ti)
                        break;
                }

                if (index < tiArray.Count() && index != searchIndex)
                {
                    searchIndex = index;
                    searchTextBox.ShowSearch = true;
                }
            }
        }

        private void BuildTreeItemList(List<TreeItem> tiList, TreeItem ti)
        {
            if (ti.Children.Count() == 0)
            {
                tiList.Add(ti);
            }
            else
            {
                foreach (var c in ti.Children)
                    BuildTreeItemList(tiList, c);
            }
        }

        private bool PropertyHitTest (int index, string text, bool system, bool display )
        {
            SystemProperty sp = tiArray[index].Item as SystemProperty;
            if ((system && sp.FullName.IndexOf(text, StringComparison.CurrentCultureIgnoreCase) >= 0) ||
                (display && sp.DisplayName != null && sp.DisplayName.IndexOf(text, StringComparison.CurrentCultureIgnoreCase) >= 0))
            {
                searchIndex = index;
                SelectTreeProperty(tiArray[index]);
                return true;
            }
            else
                return false;
        }

        private void SelectTreeProperty (TreeItem ti)
        {
            ti.IsSelected = true;

            // This is the slow part, as it involves a walk of the visual items behind the tree view
            // Unfortunately, there is no binding trick to get around it.
            var tvi = GetTreeViewItem(treeAllProperties, ti);
            if (tvi != null)
                tvi.BringIntoView();
        }

        // Microsoft provided code to search for a TreeViewItem for a given item
        // from docs.microsoft.com/en-us/dotnet/framework/wpf/controls/how-to-find-a-treeviewitem-in-a-treeview
        #region TreeViewItem search

        /// <summary>
        /// Recursively search for an item in this subtree.
        /// </summary>
        /// <param name="container">
        /// The parent ItemsControl. This can be a TreeView or a TreeViewItem.
        /// </param>
        /// <param name="item">
        /// The item to search for.
        /// </param>
        /// <returns>
        /// The TreeViewItem that contains the specified item.
        /// </returns>
        private TreeViewItem GetTreeViewItem(ItemsControl container, object item)
        {
            if (container != null)
            {
                if (container.DataContext == item)
                {
                    return container as TreeViewItem;
                }

                // Expand the current container
                if (container is TreeViewItem && !((TreeViewItem)container).IsExpanded)
                {
                    container.SetValue(TreeViewItem.IsExpandedProperty, true);
                }

                // Try to generate the ItemsPresenter and the ItemsPanel.
                // by calling ApplyTemplate.  Note that in the 
                // virtualizing case even if the item is marked 
                // expanded we still need to do this step in order to 
                // regenerate the visuals because they may have been virtualized away.

                container.ApplyTemplate();
                ItemsPresenter itemsPresenter =
                    (ItemsPresenter)container.Template.FindName("ItemsHost", container);
                if (itemsPresenter != null)
                {
                    itemsPresenter.ApplyTemplate();
                }
                else
                {
                    // The Tree template has not named the ItemsPresenter, 
                    // so walk the descendents and find the child.
                    itemsPresenter = FindVisualChild<ItemsPresenter>(container);
                    if (itemsPresenter == null)
                    {
                        container.UpdateLayout();

                        itemsPresenter = FindVisualChild<ItemsPresenter>(container);
                    }
                }

                Panel itemsHostPanel = (Panel)VisualTreeHelper.GetChild(itemsPresenter, 0);


                // Ensure that the generator for this panel has been created.
                UIElementCollection children = itemsHostPanel.Children;

                MyVirtualizingStackPanel virtualizingPanel =
                    itemsHostPanel as MyVirtualizingStackPanel;

                for (int i = 0, count = container.Items.Count; i < count; i++)
                {
                    TreeViewItem subContainer;
                    if (virtualizingPanel != null)
                    {
                        // Bring the item into view so 
                        // that the container will be generated.
                        virtualizingPanel.BringIntoView(i);

                        subContainer =
                            (TreeViewItem)container.ItemContainerGenerator.
                            ContainerFromIndex(i);
                    }
                    else
                    {
                        subContainer =
                            (TreeViewItem)container.ItemContainerGenerator.
                            ContainerFromIndex(i);

                        // Bring the item into view to maintain the 
                        // same behavior as with a virtualizing panel.
                        subContainer.BringIntoView();
                    }

                    if (subContainer != null)
                    {
                        // Search the next level for the object.
                        TreeViewItem resultContainer = GetTreeViewItem(subContainer, item);
                        if (resultContainer != null)
                        {
                            return resultContainer;
                        }
                        else
                        {
                            // The object is not under this TreeViewItem
                            // so collapse it.
                            subContainer.IsExpanded = false;
                        }
                    }
                }
            }

            return null;
        }

        /// <summary>
        /// Search for an element of a certain type in the visual tree.
        /// </summary>
        /// <typeparam name="T">The type of element to find.</typeparam>
        /// <param name="visual">The parent element.</param>
        /// <returns></returns>
        private T FindVisualChild<T>(Visual visual) where T : Visual
        {
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(visual); i++)
            {
                Visual child = (Visual)VisualTreeHelper.GetChild(visual, i);
                if (child != null)
                {
                    T correctlyTyped = child as T;
                    if (correctlyTyped != null)
                    {
                        return correctlyTyped;
                    }

                    T descendent = FindVisualChild<T>(child);
                    if (descendent != null)
                    {
                        return descendent;
                    }
                }
            }

            return null;
        }
    }

    public class MyVirtualizingStackPanel : VirtualizingStackPanel
    {
        /// <summary>
        /// Publically expose BringIndexIntoView.
        /// </summary>
        public void BringIntoView(int index)
        {
            this.BringIndexIntoView(index);
        }
    }
#endregion
}
