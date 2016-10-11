// Copyright (c) 2015, Dijji, and released under Ms-PL.  This, with other relevant licenses, can be found in the root of this distribution.

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Windows;
using AssociationMessages;

namespace FileMetadataAssociationManager
{
    class MainView : INotifyPropertyChanged
    {
        private State state;
        private MainWindow window;
        private Profile selectedProfile = null;
        private List<Extension> selectedExtensions = new List<Extension>();
        private enum HandlerSet
        {
            None,
            Ours,
            Foreign,
            Other,
        }
        private HandlerSet? handlersSelected;
        private bool sortRequired = false;
        private Profile customPreviewProfileBase = null;  // Baseline profile for selected foreign handler
        private Profile customPreviewProfile = null;      // Baseline profile merged with selected profile
        private static Profile noProfile = new Profile { Name = LocalizedMessages.NullProfile, IsNull = true };  // Dummy profile to allow selection of <None>

        public ObservableCollectionWithReset<Extension> Extensions { get { return state.Extensions; } }
        public string Restrictions { get { return state.Restrictions; } }
        public int RestrictionLevel { get { return state.RestrictionLevel; } }

        public ObservableCollection<TreeItem> FullDetails { get { return selectedProfile == null ? null : selectedProfile.FullDetails; } }
        public ObservableCollection<PropertyListEntry> PreviewDetails { get { return selectedProfile == null ? null : selectedProfile.PreviewDetails; } }

        public List<Extension> SelectedExtensions { get { return selectedExtensions; } }

        public Profile ProfileToDisplay { get { return customPreviewProfile ?? SelectedProfile; } }

        public Profile SelectedProfile
        {
            get
            {
                if (selectedProfile == null)
                    return noProfile;
                else
                    return selectedProfile;
            }
            set
            {
                Profile newProfile = value;

                if (newProfile == noProfile)
                    newProfile = null;

                if (selectedProfile != newProfile)
                {
                    selectedProfile = newProfile;
                    if (customPreviewProfileBase != null)
                    {
                        if (selectedProfile != null)
                        {
                            customPreviewProfile = customPreviewProfileBase.CreateClone();
                            customPreviewProfile.MergeFrom(selectedProfile);
                        }
                        else
                            customPreviewProfile = customPreviewProfileBase;
                    }
                    OnPropertyChanged("SelectedProfile");
                    OnPropertyChanged("ProfileToDisplay");
                    OnPropertyChanged("FullDetails");
                    OnPropertyChanged("PreviewDetails");
                    OnPropertyChanged("InfoTips");
                }
            }
        }

        public IEnumerable<Profile> Profiles
        {
            get
            {
                yield return noProfile;
                foreach (Profile p in state.BuiltInProfiles)
                {
                    yield return p;
                }
                foreach (Profile p in state.CustomProfiles)
                {
                    yield return p;
                }
            }
        }

        public bool CanChooseProfile { get { return handlersSelected == HandlerSet.None || handlersSelected == HandlerSet.Foreign; } }

        public bool CanAddPropertyHandlerEtc
        {
            get
            {
                return Extension.IsOurPropertyHandlerRegistered && SelectedProfile != null && 
                       (handlersSelected == HandlerSet.None || handlersSelected == HandlerSet.Foreign);
            }
        }

        public bool CanRemovePropertyHandlerEtc { get { return handlersSelected == HandlerSet.Ours; } }

        public bool SortRequired
        {
            get { return sortRequired; }
            set
            {
                if (value != sortRequired)
                {
                    sortRequired = value;
                    OnPropertyChanged("SortRequired");
                }
            }
        }
        
        public MainView(MainWindow window, State state)
        {
            this.window = window;
            this.state = state;
        }

        public void SetSelectedExtensions(System.Collections.IList selections)
        {
            SelectedExtensions.Clear();
            foreach (var s in selections)
                SelectedExtensions.Add((Extension)s);

            DeterminePossibleActions();
        }

        public void AddHandlers()
        {
            if (SelectedExtensions.Count > 0 && SelectedProfile != null)
            {
                if (SelectedExtensions.First().PropertyHandlerState == HandlerState.None && SelectedProfile.IsNull)
                {
                    MessageBox.Show(LocalizedMessages.PleaseSelectProfile, LocalizedMessages.SetupHandler);
                    return;
                }
                else if (SelectedExtensions.First().PropertyHandlerState == HandlerState.Foreign)
                {
                    if (SelectedProfile.IsNull)
                    {
                        if (MessageBox.Show(SelectedExtensions.Count > 1 ? LocalizedMessages.ConfirmCustomNoMerges : LocalizedMessages.ConfirmCustomNoMerge, 
                            LocalizedMessages.SetupHandler, MessageBoxButton.YesNo) == MessageBoxResult.No)
                            return;
                    }
                    else
                    {
                        if (MessageBox.Show(string.Format(SelectedExtensions.Count > 1 ? LocalizedMessages.ConfirmCustomMerges : LocalizedMessages.ConfirmCustomMerge, 
                            SelectedProfile.Name), LocalizedMessages.SetupHandler, MessageBoxButton.YesNo) == MessageBoxResult.No)
                            return;
                    }
                }

                foreach (Extension ext in SelectedExtensions)
                {
                    ext.SetupHandlerForExtension(SelectedProfile, true);
                }
            }

            OnPropertyChanged("Profiles");
            DeterminePossibleActions();
            SortRequired = true;
        }

        public void RemoveHandlers()
        {
            foreach (Extension ext in SelectedExtensions)
            {
                ext.RemoveHandlerFromExtension();
            }

            DeterminePossibleActions();
            SortRequired = true;
        }

        public void RefreshProfiles()
        {
            OnPropertyChanged("Profiles");
        }

        public void RefreshRegistry()
        {
            int count = 0;
            foreach (var ext in state.Extensions.Where(e => e.PropertyHandlerState == HandlerState.Ours ||  e.PropertyHandlerState == HandlerState.Chained))
            {
                if (ext.IsRefreshRequired())
                {
                    var p = ext.Profile;
                    ext.RemoveHandlerFromExtension();
                    ext.SetupHandlerForExtension(p, true);
                    count++;
                }
            }

            if (count == 0)
                MessageBox.Show(LocalizedMessages.NoRegistryUpdatesNeeded, LocalizedMessages.RegistryUpdate);
            else
                MessageBox.Show(String.Format(LocalizedMessages.RegistryUpdatesMade, count), LocalizedMessages.RegistryUpdate);
        }

        private void DeterminePossibleActions()
        {
            // Cases are:
            // 1. All selected extensions have no handler: profile combo box is enabled and profile property lists are shown. 
            // 2. All selected extensions have File Meta handler: profile combo box is disabled and profile for the 
            //    first selected extension is shown in combo box and profile property lists. 
            // 3. All other cases: profile combo box is disabled and profile property lists are empty.
            handlersSelected = null;
            foreach (Extension e in SelectedExtensions)
            {
                if (e.PropertyHandlerState == HandlerState.None)
                {
                    if (handlersSelected == null)
                        handlersSelected = HandlerSet.None;
                    else if (handlersSelected == HandlerSet.None)
                        continue;
                    else
                    {
                        handlersSelected = HandlerSet.Other;
                        break;
                    }
                }
                else if (e.PropertyHandlerState == HandlerState.Foreign)
                {
                    if (handlersSelected == null)
                        handlersSelected = HandlerSet.Foreign;
                    else if (handlersSelected == HandlerSet.Foreign)
                        continue;
                    else
                    {
                        handlersSelected = HandlerSet.Other;
                        break;
                    }
                }
                else if (e.PropertyHandlerState == HandlerState.Ours || e.PropertyHandlerState == HandlerState.Chained)
                {
                    if (handlersSelected == null)
                        handlersSelected = HandlerSet.Ours;
                    else if (handlersSelected == HandlerSet.Ours)
                        continue;
                    else
                    {
                        handlersSelected = HandlerSet.Other;
                        break;
                    }
                }
                else
                {
                    handlersSelected = HandlerSet.Other;
                    break;
                }
            }
            if (handlersSelected == null)
                handlersSelected = HandlerSet.Other;

            customPreviewProfile = null;
            customPreviewProfileBase = null;

            switch (handlersSelected)
            {
                case HandlerSet.Ours:
                    SelectedProfile = SelectedExtensions.First().Profile;
                    break;
                case HandlerSet.Foreign:
                {
                    if (SelectedExtensions.Count == 1)
                    {
                        customPreviewProfileBase = SelectedExtensions.First().GetDefaultCustomProfile();
                        customPreviewProfile = customPreviewProfileBase;
                    }

                    SelectedProfile = null;
                    break;
                }
                case HandlerSet.None:
                case HandlerSet.Other:
                    SelectedProfile = null;
                    break;
            }

            OnPropertyChanged("ProfileToDisplay");
            OnPropertyChanged("CanChooseProfile");
            OnPropertyChanged("CanAddPropertyHandlerEtc");
            OnPropertyChanged("CanRemovePropertyHandlerEtc");
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
