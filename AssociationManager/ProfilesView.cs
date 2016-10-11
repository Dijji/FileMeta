// Copyright (c) 2015, Dijji, and released under Ms-PL.  This, with other relevant licenses, can be found in the root of this distribution.

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using AssociationMessages;
using GongSolutions.Wpf.DragDrop;

namespace FileMetadataAssociationManager
{
    class ProfilesView : INotifyPropertyChanged, IDropTarget
    {
        private State state;
        private ProfilesWindow window;
        private List<Profile> editableProfiles = new List<Profile>();
        private List<Profile> deletedProfiles = new List<Profile>();
        private Profile selectedProfile = null;
        private bool isInTreeTextBoxEditMode = false;
        private bool isDirty = false;

        public List<string> GroupProperties { get { return state.GroupProperties; } }
        public List<TreeItem> AllProperties { get { return state.AllProperties; } }

        public ObservableCollection<TreeItem> FullDetails { get { return selectedProfile == null ? null : selectedProfile.FullDetails; } }
        public ObservableCollection<PropertyListEntry> PreviewDetails { get { return selectedProfile == null ? null : selectedProfile.PreviewDetails; } }
        public ObservableCollection<PropertyListEntry> InfoTips { get { return selectedProfile == null ? null : selectedProfile.InfoTips; } }

        public Profile SelectedProfile 
        {
            get { return selectedProfile; } 
            set
            {
                if (selectedProfile != value)
                {
                    selectedProfile = value;
                    OnPropertyChanged("SelectedProfile");
                    OnPropertyChanged("FullDetails");
                    OnPropertyChanged("PreviewDetails");
                    OnPropertyChanged("InfoTips");
                }
            }
        }

        public bool IsSelectedProfileWritable { get { return IsProfileWritable (SelectedProfile); } }

        public bool IsInTreeTextBoxEditMode
        {
            get { return isInTreeTextBoxEditMode; }
            set
            {
                if (value != isInTreeTextBoxEditMode)
                {
                    isInTreeTextBoxEditMode = value;
                    OnPropertyChanged("IsInTreeTextBoxEditMode");
                }
            }
        }

        public bool IsDirty
        {
            get { return isDirty; }
            private set
            {
                if (value != isDirty)
                {
                    isDirty = value;
                    OnPropertyChanged("IsDirty");
                }
            }
        }

        public List<TreeItem> ProfileTree
        {
            get
            {
                // Return a tree describing the profiles
                List<TreeItem> profileTree = new List<TreeItem>();
                TreeItem ti = new TreeItem(LocalizedMessages.ProfileBuiltIn);

                // The first leg contains the built-in profiles
                ProfileTreeHelper(state.BuiltInProfiles, ti);
                profileTree.Add(ti);

                // The second leg contains the custom profiles
                ti = new TreeItem(LocalizedMessages.ProfileCustom);
                ProfileTreeHelper(editableProfiles, ti);
                ti.Children.Sort((a, b) => a.Name.CompareTo(b.Name));
                profileTree.Add(ti);

                return profileTree;
            }
        }

        // Helper for building a tree describing the profiles
        private void ProfileTreeHelper(IEnumerable<Profile> profiles, TreeItem parent)
        {
            foreach (Profile p in profiles)
            {
                var ti = new TreeItem(p.Name, p);

                // Hook on a handler to deal with name changes
                ti.NameChanged += new NameChangedEventHandler(treeItem_NameChanged);

                // Selection is controlled programmatically through binding to IsSelected
                if (p == SelectedProfile)
                    ti.IsSelected = true;

                parent.AddChild(ti);
            }
        }

        // Constructor
        public ProfilesView(ProfilesWindow window, State state)
        {
            this.window = window;
            this.state = state;

            //  Make a working copy of the custom profiles so that we can discard edits later, if we are cancelled
            foreach(var p in state.CustomProfiles)
            {
                var editableProfile = p.CreateClone();
                editableProfile.Original = p;
                editableProfiles.Add(editableProfile);
                if (p == state.SelectedProfile)
                    SelectedProfile = editableProfile;
            }

            // Reflect the profile selected in the main window, if possible
            if (state.SelectedProfile != null && state.SelectedProfile.IsReadOnly)
                SelectedProfile = state.SelectedProfile;
        }

        public bool IsProfileWritable (Profile p)
        {
            return p != null && p.IsReadOnly == false;
        }  

        public bool ApplyChanges()
        {
            // Only do anything if there have been some changes
            if (IsDirty)
            {
                // As a first pass, see if any changes impact any extensions configured with our handler
                string impacted = "";
                foreach (var p in editableProfiles)
                {
                    // Look at profiles that have not been added to see if they have been changed
                    if (p.Original != null && p.DiffersFrom(p.Original))
                    {
                        foreach (var ext in state.GetExtensionsUsingProfile(p.Original))
                        {
                           // Build a running list of affected extensions
                           if (impacted.Length > 0)
                               impacted += ", ";
                           impacted += ext.Name;
                        }
                    }
                }

                // If there were any impacted extensions, give the user a chance to back out
                if (impacted.Length > 0)
                {
                    var result = MessageBox.Show(String.Format(LocalizedMessages.ChangingProfilesQuestion, impacted),
                                                 LocalizedMessages.ChangingProfilesHeader, MessageBoxButton.YesNo);

                    if (result == MessageBoxResult.No)
                        return false;

                    state.HasChanged = true;
                }

                // Proceed with applying all the changes
                foreach (var p in editableProfiles)
                {
                    // Added profiles will not have an original
                    if (p.Original == null)
                    {
                        state.CustomProfiles.Add(p);
                    }
                    else if (p.DiffersFrom(p.Original))
                    {
                        // Update the registry settings
                        foreach (var ext in state.GetExtensionsUsingProfile(p.Original))
                        {
                            ext.UpdateProfileSettingsForExtension(p);
                        }
                        // And then the profile of record
                        p.Original.UpdateFrom(p);
                    }
                }

                // Finally, remove any deleted profiles
                foreach (var p in deletedProfiles)
                {
                    if (p.Original != null)
                        state.CustomProfiles.Remove(p.Original);
                }

                // Write the updated custom profile information back to the store
                state.StoreSavedState();
                IsDirty = false;
            }

            return true;
        }

        // Handle a profile name change
        void treeItem_NameChanged(object sender, NameChangedEventArgs e)
        {
            TreeItem ti = sender as TreeItem;

            // Don't allow reuse of an existing name
            if (ExistingProfileNamed(e.NewName))
            {
                MessageBox.Show(String.Format(LocalizedMessages.ProfileAlreadyExists, e.NewName), LocalizedMessages.ProfileError);
                ti.AbandonNameChange();
                return;
            }

            Profile p = ti.Item as Profile;
            if (p != null)
            {
                ti.ChangeName(e.NewName);
                p.Name = e.NewName;
                IsDirty = true;
            }
        }

        // Add a new profile, which may be empty, or copied from an existing profile
        public Profile AddNewProfile(Profile copyFrom = null)
        {
            Profile p;
            if (copyFrom == null)
                p = new Profile();
            else
                p = copyFrom.CreateClone();

            // Generate a name for the new profile, adding (N) as necessary for uniqueness
            int index = 1;
            string name = Profile.NewName(index);
            while (ExistingProfileNamed(name))
            {
                name = Profile.NewName(++index);
            }
            p.Name = name;
            editableProfiles.Add(p);

            // Select the new profile, and tell the profile tree to update
            IsDirty = true;
            SelectedProfile = p;
            OnPropertyChanged("ProfileTree");

            return p;
        }

        public bool ExistingProfileNamed(string name)
        {
            return state.BuiltInProfiles.FirstOrDefault(p => string.Equals(p.Name, name, StringComparison.CurrentCultureIgnoreCase)) != null ||
                   editableProfiles.FirstOrDefault(p => string.Equals(p.Name, name, StringComparison.CurrentCultureIgnoreCase)) != null;
        }

        public void DeleteProfile(Profile p)
        {
            // Cannot delete a profile that is in use by one or more extensions
            if (p.Original != null && state.GetExtensionsUsingProfile(p.Original).Count() > 0)
            {
                MessageBox.Show(String.Format(LocalizedMessages.CannotDeleteProfile, p.Name), LocalizedMessages.ProfileError);
                return;
            }

            int index = editableProfiles.IndexOf(p);

            if (index >= 0)
            {
                // If we're deleting the selected profile, try and move the selection to the one before it
                if (SelectedProfile == p)
                {
                    SelectedProfile = index > 0 ? editableProfiles[index - 1] : editableProfiles.FirstOrDefault();
                }

                // Move the profile to the deleted list, so that we can apply the change to the underlying state, later
                editableProfiles.RemoveAt(index);
                deletedProfiles.Add(p);
                IsDirty = true;
                OnPropertyChanged("ProfileTree");
            }
        }
        
        public void RemoveFullDetailsItem(TreeItem ti)
        {
            if (IsSelectedProfileWritable)
            {
                // Prompt the user for confirmation in the cases where we are removing a group, or
                // the property also appears in the preview details, and so will be removed there too
                MessageBoxResult ans = MessageBoxResult.Yes;
                if ((PropType)ti.Item == PropType.Group)
                    ans = MessageBox.Show(String.Format(LocalizedMessages.RemovePropertyGroupQuestion, ti.Name),
                                          LocalizedMessages.RemovePropertyGroupHeader, MessageBoxButton.YesNo);
                else if (SelectedProfile.HasPropertyInPreviewDetails(ti.Name) || SelectedProfile.HasPropertyInInfoTip(ti.Name))
                    ans = MessageBox.Show(String.Format(LocalizedMessages.RemovePropertyQuestion, ti.Name),
                                       LocalizedMessages.RemovePropertyHeader, MessageBoxButton.YesNo);

                if (ans == MessageBoxResult.Yes)
                {
                    SelectedProfile.RemoveFullDetailsItem(ti);
                    IsDirty = true;
                }
            }
        }

        public void ToggleAsteriskFullDetailsItem(TreeItem ti)
        {
            if (IsSelectedProfileWritable)
            {
                SelectedProfile.ToggleAsteriskFullDetailsItem(ti);
                IsDirty = true;
            }
        }

        public void RemovePreviewDetailsItem(PropertyListEntry property)
        {
            if (IsSelectedProfileWritable)
            {
                SelectedProfile.RemovePreviewDetailsProperty(property);
                IsDirty = true;
            }
        }
        
        public void RemoveInfoTipItem(PropertyListEntry property)
        {
            if (IsSelectedProfileWritable)
            {
                SelectedProfile.RemoveInfoTipProperty(property);
                IsDirty = true;
            }
        }

        public void ToggleAsterisk(PropertyListEntry property)
        {
            if (IsSelectedProfileWritable)
            {
                property.ToggleAsterisk();
                IsDirty = true;
            }
        }

        // This is invoked whenever and item is dragged over one of our drop targets.
        // The main job is to determine whether this is a viable drop.
        // If we say nothing, the default answer is no.
        public void DragOver(IDropInfo dropInfo)
        {
            // Drops are only valid on to writable profiles
            if (IsSelectedProfileWritable)
            {
                var Source = window.IdentifyControl(dropInfo.DragInfo.VisualSource);
                var Target = window.IdentifyControl(dropInfo.VisualTarget);

                // Dropping a group onto Full Details 
                if (Source == ProfileControls.Groups && Target == ProfileControls.FullDetails)
                {
                    // Check group is not already there
                    string group = dropInfo.Data as string;

                    if (group != null && !SelectedProfile.HasGroupInFullDetails(group))
                    {
                        dropInfo.DropTargetAdorner = DropTargetAdorners.Insert;
                        dropInfo.Effects = DragDropEffects.Copy;
                    }
                }

                // Dropping a property onto Full Details
                else if (Source == ProfileControls.Properties && Target == ProfileControls.FullDetails)
                {
                    //System.Diagnostics.Trace.WriteLine(String.Format("DragOver {0} at {1}", ((TreeItem)dropInfo.Data).Name,
                    //                                   dropInfo.TargetItem == null ? "null" : ((TreeItem)dropInfo.TargetItem).Name));

                    // Check at least one group is present
                    if (SelectedProfile.FullDetails.Count > 0)
                    {
                        // Check property is not already there
                        var ti = dropInfo.Data as TreeItem;
                        string property = state.GetSystemPropertyName(ti);

                        if (property != null && !SelectedProfile.HasPropertyInFullDetails(property))
                        {
                            dropInfo.DropTargetAdorner = DropTargetAdorners.Insert;
                            dropInfo.Effects = DragDropEffects.Copy;
                        }
                    }
                }

                // Reordering within Full Details
                else if (Source == ProfileControls.FullDetails && Target == ProfileControls.FullDetails)
                {
                    //System.Diagnostics.Trace.WriteLine(String.Format("DragOver {0} at {1}", ((TreeItem)dropInfo.Data).Name,
                    //               dropInfo.TargetItem == null ? "null" : ((TreeItem)dropInfo.TargetItem).Name));

                    var ti = dropInfo.Data as TreeItem;

                    if (ti != null)
                    {
                        dropInfo.DropTargetAdorner = DropTargetAdorners.Insert;
                        dropInfo.Effects = DragDropEffects.Move;
                    }
                }

                // Dragging from Full Details to Preview Details
                else if (Source == ProfileControls.FullDetails && Target == ProfileControls.PreviewDetails)
                {
                    var ti = dropInfo.Data as TreeItem;

                    // Can only drag properties, not groups, and only if they are not already present
                    if (ti != null &&
                        (PropType)ti.Item == PropType.Normal &&
                        !SelectedProfile.HasPropertyInPreviewDetails(ti.Name))
                    {
                        dropInfo.DropTargetAdorner = DropTargetAdorners.Insert;
                        dropInfo.Effects = DragDropEffects.Copy;
                    }
                }

                // Reordering within Preview Details
                else if (Source == ProfileControls.PreviewDetails && Target == ProfileControls.PreviewDetails)
                {
                    var property = dropInfo.Data as PropertyListEntry;

                    if (property != null)
                    {
                        dropInfo.DropTargetAdorner = DropTargetAdorners.Insert;
                        dropInfo.Effects = DragDropEffects.Move;
                    }
                }

                // Dragging from Full Details to InfoTip
                else if (Source == ProfileControls.FullDetails && Target == ProfileControls.InfoTip)
                {
                    var ti = dropInfo.Data as TreeItem;

                    // Can only drag properties, not groups, and only if they are not already present
                    if (ti != null &&
                        (PropType)ti.Item == PropType.Normal &&
                        !SelectedProfile.HasPropertyInInfoTip(ti.Name))
                    {
                        dropInfo.DropTargetAdorner = DropTargetAdorners.Insert;
                        dropInfo.Effects = DragDropEffects.Copy;
                    }
                }

                // Reordering within InfoTip
                else if (Source == ProfileControls.InfoTip && Target == ProfileControls.InfoTip)
                {
                    var property = dropInfo.Data as PropertyListEntry;

                    if (property != null)
                    {
                        dropInfo.DropTargetAdorner = DropTargetAdorners.Insert;
                        dropInfo.Effects = DragDropEffects.Move;
                    }
                }
            }
        }

        // Handle an actual drop.
        // We don't have to do so many checks here, because we have already validated the drop in DragOver
        public void Drop(IDropInfo dropInfo)
        {
            var Source = window.IdentifyControl(dropInfo.DragInfo.VisualSource);
            var Target = window.IdentifyControl(dropInfo.VisualTarget);

            if (Target == ProfileControls.FullDetails)
            {
                TreeItem target = dropInfo.TargetItem as TreeItem;
                bool before = (dropInfo.InsertPosition & ~RelativeInsertPosition.TargetItemCenter) == RelativeInsertPosition.BeforeTargetItem;

                if (Source == ProfileControls.Groups)
                {
                    string group = dropInfo.Data as string;

                    SelectedProfile.AddFullDetailsGroup(group, target, before);
                    IsDirty = true;
                }
                else if (Source == ProfileControls.Properties)
                {
                    //System.Diagnostics.Trace.WriteLine(String.Format("Drop {0} {1} {2}", ((TreeItem)dropInfo.Data).Name,
                    //                                   dropInfo.InsertPosition & ~RelativeInsertPosition.TargetItemCenter,
                    //                                   dropInfo.TargetItem == null ? "null" : ((TreeItem)dropInfo.TargetItem).Name));

                    string property = state.GetSystemPropertyName((TreeItem)dropInfo.Data);

                    SelectedProfile.AddFullDetailsProperty(property, target, before);
                    IsDirty = true;
                }
                else if (Source == ProfileControls.FullDetails)
                {
                    //System.Diagnostics.Trace.WriteLine(String.Format("Drop {0} {1} {2}", ((TreeItem)dropInfo.Data).Name,
                    //                                   dropInfo.InsertPosition & ~RelativeInsertPosition.TargetItemCenter,
                    //                                   dropInfo.TargetItem == null ? "null" : ((TreeItem)dropInfo.TargetItem).Name));

                    TreeItem toMove = (TreeItem)dropInfo.Data;

                    // Group items are distinguishable from property items by means of the PropType value stored on them
                    if ((PropType)toMove.Item == PropType.Group)
                        SelectedProfile.MoveFullDetailsGroup(toMove, target, before);
                    else
                        SelectedProfile.MoveFullDetailsProperty(toMove, target, before);

                    IsDirty = true;
                    toMove.IsSelected = true;  //todo make work
                }
            }
            else if (Target == ProfileControls.PreviewDetails)
            {
                PropertyListEntry target = dropInfo.TargetItem as PropertyListEntry;
                bool before = (dropInfo.InsertPosition & ~RelativeInsertPosition.TargetItemCenter) == RelativeInsertPosition.BeforeTargetItem;

                if (Source == ProfileControls.FullDetails)
                {
                    string property = ((TreeItem)dropInfo.Data).Name;
                    SelectedProfile.AddPreviewDetailsProperty(property, target, before);
                }
                else if (Source == ProfileControls.PreviewDetails)
                {
                    PropertyListEntry property = (PropertyListEntry)dropInfo.Data;
                    SelectedProfile.MovePreviewDetailsProperty(property, target, before);
                }
                IsDirty = true;
            }
            else if (Target == ProfileControls.InfoTip)
            {
                PropertyListEntry target = dropInfo.TargetItem as PropertyListEntry;
                bool before = (dropInfo.InsertPosition & ~RelativeInsertPosition.TargetItemCenter) == RelativeInsertPosition.BeforeTargetItem;

                if (Source == ProfileControls.FullDetails)
                {
                    string property = ((TreeItem)dropInfo.Data).Name;
                    SelectedProfile.AddInfoTipProperty(property, target, before);
                }
                else if (Source == ProfileControls.InfoTip)
                {
                    PropertyListEntry property = (PropertyListEntry)dropInfo.Data;
                    SelectedProfile.MoveInfoTipProperty(property, target, before);
                }
                IsDirty = true;
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
