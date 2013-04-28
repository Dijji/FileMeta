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
        private Extension selectedExtension = null;
        private Profile selectedProfile = null;
        private ObservableCollection<TreeItem> allProperties = new ObservableCollection<TreeItem>();
        private ObservableCollection<Profile> profiles = new ObservableCollection<Profile>();
        private List<string> groupProperties = new List<string>();

        public ObservableCollectionWithReset<Extension> Extensions { get { return extensions; } }
        public ObservableCollection<TreeItem> AllProperties { get { return allProperties; } }
        public ObservableCollection<Profile> Profiles { get { return profiles; } }

        public List<string> GroupProperties { get { return groupProperties; } }
        public Extension SelectedExtension { get { return selectedExtension; } set { selectedExtension = value; OnPropertyChanged("SelectedExtension"); } }
        public Profile SelectedProfile { get { return selectedProfile; } set { selectedProfile = value; OnPropertyChanged("SelectedProfile"); } }

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
            PopulateSystemProperties();
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
                    handlerTitle = (string)cls.GetValue(null);

                if (dictExtensions.TryGetValue(name.ToLower(), out e))
                {
                    e.RecordPropertyHandler(handlerGuid, handlerTitle);
                }
            }

            SortExtensions();
            SelectedExtension = Extensions.First();
        }

        public void SortExtensions()
        {
            // Sort by file extension, but group by our handler, other handlers, and finally no handler
            List<Extension> es = extensions.ToList();
            es.Sort((e, f) =>
                e.OurHandler ?
                    (f.OurHandler ? e.Name.CompareTo(f.Name) : -1) :
                    f.OurHandler ? 1 :
                        e.ForeignHandler ? (f.ForeignHandler ? e.Name.CompareTo(f.Name) : -1) :
                        f.ForeignHandler ? 1 :
                            e.Name.CompareTo(f.Name));
            extensions.Clear();
            foreach (var e in es)
                extensions.Add(e);
            Extensions.NotifyReset();
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

                        string propName;
                        propertyDescription.GetCanonicalName(out propName);

                        List<string> names = null;
                        string[] parts = propName.Split('.');
                        if (parts.Count() == 2)
                        {
                            // System.Foo
                            if (!dict.TryGetValue(parts[0], out names))
                            {
                                names = new List<string>();
                                dict.Add(parts[0], names);
                            }
                            names.Add(parts[1]);
                        }
                        else if (parts.Count() == 3)
                        {
                            // System.Bar.Baz
                            if (!dict.TryGetValue(parts[1], out names))
                            {
                                names = new List<string>();
                                dict.Add(parts[1], names);
                            }
                            names.Add(parts[2]);
                        }

                        // If we ever need it:
                        // ShellPropertyDescription desc = new ShellPropertyDescription(propertyDescription);

                        if (propertyDescription != null)
                        {
                            Marshal.ReleaseComObject(propertyDescription);
                            propertyDescription = null;
                        }
                    }

                    // build tree
                    foreach (string cat in dict.Keys)
                    {
                        TreeItem main = new TreeItem(cat);
                        foreach (string name in dict[cat])
                            main.Children.Add(new TreeItem(name));

                        if (cat == "System")
                            AllProperties.Insert(0, main);
                        else if (cat == "PropGroup")
                        {
                            foreach (TreeItem ti in main.Children)
                                GroupProperties.Add(ti.Name);
                        }
                        else
                            AllProperties.Add(main);
                    }
                }
            }
            finally
            {
                if (propertyDescriptionList != null)
                {
                    Marshal.ReleaseComObject(propertyDescriptionList);
                }
                if (propertyDescription != null)
                {
                    Marshal.ReleaseComObject(propertyDescription);
                }
            }
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
