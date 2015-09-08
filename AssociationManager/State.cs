// Copyright (c) 2013, Dijji, and released under Ms-PL.  This, with other relevant licenses, can be found in the root of this distribution.

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Serialization;
using Microsoft.Win32;
using System.Runtime.InteropServices;
using FileMetadataAssociationManager.Resources;

namespace FileMetadataAssociationManager
{
    public class State
    {
        private ObservableCollectionWithReset<Extension> extensions = new ObservableCollectionWithReset<Extension>();
        private Dictionary<string, Extension> dictExtensions = new Dictionary<string, Extension>();
        private List<Profile> builtInProfiles = new List<Profile>();
        private List<TreeItem> allProperties = new List<TreeItem>();
        private List<string> groupProperties = new List<string>();
        private SavedState savedState = new SavedState();
        private bool hasChanged = false;

        public ObservableCollectionWithReset<Extension> Extensions { get { return extensions; } }
        public List<Profile> BuiltInProfiles { get { return builtInProfiles; } }
        public List<Profile> CustomProfiles { get { return savedState.CustomProfiles; } }
        public List<TreeItem> AllProperties { get { return allProperties; } }
        public List<string> GroupProperties { get { return groupProperties; } }
        public bool HasChanged { get { return hasChanged; } set { hasChanged = value; } }

        public Profile SelectedProfile { get; set; }  // used during view transitions

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
            PopulateProfiles();
            PopulateExtensions();
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


        public void PopulateSystemProperties()
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
                        TreeItem main = new TreeItem(cat, PropType.Group);
                        foreach (string name in dict[cat])
                            main.AddChild(new TreeItem(name, PropType.Normal));

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

        public string GetSystemPropertyName(TreeItem ti)
        {
            if (ti == null || (PropType)ti.Item != PropType.Normal)
                return null;
            else 
            {
                StringBuilder sb = new StringBuilder(ti.Name);
                bool hasSystem = false;
                for (TreeItem parent = ti.Parent; parent != null; parent = parent.Parent)
                {
                    sb.Insert(0, ".");
                    sb.Insert(0, parent.Name);
                    if(parent.Name == "System")
                        hasSystem = true;
                }
                if (!hasSystem)
                    sb.Insert(0, "System.");

                return sb.ToString();
            }
        }

        public void LoadSavedState()
        {
            try
            {
                DirectoryInfo di = ObtainDataDirectory();
                FileInfo fi = di.GetFiles().Where(f => f.Name == "SavedState.xml").FirstOrDefault();

                if (fi != null)
                {
                    XmlSerializer x = new XmlSerializer(typeof(SavedState));
                    TextReader reader = new StreamReader(fi.FullName);
                    SavedState loaded = (SavedState)x.Deserialize(reader);
                    reader.Close();
                    savedState = loaded;
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(String.Format(LocalizedMessages.XmlParseError, ex.Message), LocalizedMessages.ErrorHeader);
            }
        }

        public void StoreSavedState()
        {
            try
            {
                XmlSerializer x = new XmlSerializer(typeof(SavedState));
                TextWriter writer = new StreamWriter(ObtainDataDirectory().FullName + @"\SavedState.xml");
                x.Serialize(writer, savedState);
                writer.Close();
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(String.Format(LocalizedMessages.XmlWriteError, ex.Message), LocalizedMessages.ErrorHeader);
            }
        }

        public IEnumerable<Extension> GetExtensionsUsingProfile(Profile p)
        {
            return Extensions.Where(e => e.Profile == p);
        }

        private void PopulateProfiles()
        {
            foreach (Profile p in Profile.GetBuiltinProfiles(this))
                BuiltInProfiles.Add(p);

            LoadSavedState();
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
                    e.IdentifyCurrentProfile();
                }
            }

            SortExtensions();
        }


        private DirectoryInfo ObtainDataDirectory()
        {
            DirectoryInfo di = new DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData));
            DirectoryInfo target = di.GetDirectories().Where(d => d.Name == "FileMeta").FirstOrDefault();
            if (target == null)
                target = di.CreateSubdirectory("FileMeta");

            return target;
        }
    }
}
