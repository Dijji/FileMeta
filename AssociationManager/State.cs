﻿// Copyright (c) 2013, Dijji, and released under Ms-PL.  This, with other relevant licenses, can be found in the root of this distribution.

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using Microsoft.Win32;
using System.Runtime.InteropServices;
using FileMetadataAssociationManager.Resources;

namespace FileMetadataAssociationManager
{
    public class State : INotifyPropertyChanged
    {
        private ObservableCollectionWithReset<Extension> extensions = new ObservableCollectionWithReset<Extension>();
        private Dictionary<string, Extension> dictExtensions = new Dictionary<string, Extension>();
        private List<Extension> selectedExtensions = new List<Extension>();
        private Profile selectedProfile = null;
        private ObservableCollection<Profile> profiles = new ObservableCollection<Profile>();
        private enum HandlerSet
        {
            None,
            Ours,
            Other,
        }
        private HandlerSet? selectedHandlers;

        public ObservableCollectionWithReset<Extension> Extensions { get { return extensions; } }
        public ObservableCollection<Profile> Profiles { get { return profiles; } }

        public List<Extension> SelectedExtensions { get { return selectedExtensions; } }
        public Profile SelectedProfile { get { return selectedProfile; } set { selectedProfile = value; OnPropertyChanged("SelectedProfile"); } }

        public void SetSelectedExtensions(System.Collections.IList selections)
        {
            SelectedExtensions.Clear();
            foreach (var s in selections)
                SelectedExtensions.Add((Extension)s);

            // Cases are:
            // 1. All selected extensions have no handler: profile combo box is enabled and profile property lists are shown. 
            // 2. All selected extensions have File Meta handler: profile combo box is disabled and profile for the 
            //    first selected extension is shown in combo box and profile property lists. 
            // 3. All other cases: profile combo box is disabled and profile property lists are empty.
            selectedHandlers = null;
            foreach (Extension e in SelectedExtensions)
            {
                if (!e.HasHandler)
                {
                    if (selectedHandlers == null)
                        selectedHandlers = HandlerSet.None;
                    else if (selectedHandlers == HandlerSet.None)
                        continue;
                    else
                    {
                        selectedHandlers = HandlerSet.Other;
                        break;
                    }
                }
                else if (e.OurHandler)
                {
                    if (selectedHandlers == null)
                        selectedHandlers = HandlerSet.Ours;
                    else if (selectedHandlers == HandlerSet.Ours)
                        continue;
                    else
                    {
                        selectedHandlers = HandlerSet.Other;
                        break;
                    }
                }
                else
                {
                    selectedHandlers = HandlerSet.Other;
                    break;
                }
            }
            if (selectedHandlers == null)
                selectedHandlers = HandlerSet.Other;

            switch (selectedHandlers)
            {
                case HandlerSet.None:
                    if (SelectedProfile == null)
                        SelectedProfile = Profiles.First();
                    break;
                case HandlerSet.Ours:
                    SelectedProfile = SelectedExtensions.First().GetCurrentProfileIfKnown();
                    break;
                case HandlerSet.Other:
                    SelectedProfile = null;
                    break;
            }

            OnPropertyChanged("CanChooseProfile");
            OnPropertyChanged("CanAddPropertyHandlerEtc");
            OnPropertyChanged("CanRemovePropertyHandlerEtc");
        }

        public bool CanChooseProfile { get { return selectedHandlers == HandlerSet.None; } }

        public bool CanAddPropertyHandlerEtc
        {
            get
            {
                return Extension.IsOurPropertyHandlerRegistered && SelectedProfile != null &&
                       selectedHandlers == HandlerSet.None;
            }
        }
        public bool CanRemovePropertyHandlerEtc { get { return selectedHandlers == HandlerSet.Ours; } }

        public string Restrictions
        {
            get
            {
                if (!Extension.IsOurPropertyHandlerRegistered)
                    return LocalizedMessages.CannotAddHandler;
                else if (!Extension.IsOurContextHandlerRegistered)
                    return LocalizedMessages.NoContextMenus;
                else
                    return "";
            }
        }

        public int RestrictionLevel
        {
            get
            {
                if (!Extension.IsOurPropertyHandlerRegistered)
                    return 2;
                else if (!Extension.IsOurContextHandlerRegistered)
                    return 1;
                else
                    return 0;
            }
        }

        public void Populate()
        {
            PopulateExtensions();
            PopulateProfiles();
        }
        
        private void PopulateExtensions()
        {
            var hkcr = Registry.ClassesRoot;

            // get all the extensions in the registry
            foreach (string name in hkcr.GetSubKeyNames())
            {
                if (name.StartsWith("."))
                {
                    Extension e = new Extension() { Name = name, State = this };
                    dictExtensions.Add(name.ToLower(), e);
                    Extensions.Add(e);
                }
            }

            // get all the current handlers
            var handlers = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\PropertySystem\PropertyHandlers", false);
            foreach (string name in handlers.GetSubKeyNames())
            {
                string handlerGuid = (string)handlers.OpenSubKey(name, false).GetValue(null);
                string handlerTitle = null;
                Extension e;
                var cls = hkcr.OpenSubKey(@"CLSID\" + handlerGuid);
                if (cls != null)
                {
                    handlerTitle = (string)cls.GetValue(null);
                    if (handlerTitle == null)
                    {
                        // No name - check for shell handlers
                        if (handlerGuid == "{66742402-F9B9-11D1-A202-0000F81FEDEE}" ||
                            handlerGuid == "{0AFCCBA6-BF90-4A4E-8482-0AC960981F5B}")
                            handlerTitle = "Windows Shell";
                        else
                        {
                            // Else resort to the dll path
                            var ps = cls.OpenSubKey("InProcServer32");
                            if (ps != null)
                                handlerTitle = (string)ps.GetValue(null);
                        }
                    }
                }

                if (dictExtensions.TryGetValue(name.ToLower(), out e))
                {
                    e.RecordPropertyHandler(handlerGuid, handlerTitle);
                }
            }

            SortExtensions();
            SetSelectedExtensions(new List<Extension> {Extensions.First()});
        }

        public void SortExtensions()
        {
            // Sort by file extension, but group by our handler, other handlers, and finally no handler
            // This uses a Sort extension to ObservableCollection
            extensions.Sort((e, f) =>
                e.OurHandler ?
                    (f.OurHandler ? e.Name.CompareTo(f.Name) : -1) :
                    f.OurHandler ? 1 :
                        e.ForeignHandler ? (f.ForeignHandler ? e.Name.CompareTo(f.Name) : -1) :
                        f.ForeignHandler ? 1 :
                            e.Name.CompareTo(f.Name));
            //Extensions.NotifyReset();
        }

        private void PopulateProfiles()
        {
            foreach (Profile p in Profile.GetBuiltinProfiles(this))
                Profiles.Add(p);

            SelectedProfile = Profiles.Count > 0 ? Profiles[0] : null;
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
