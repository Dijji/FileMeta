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
using AssociationMessages;

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
        private bool nonDefaultStateFileLoaded = false;
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

        public void Populate(string savedStateFile = null)
        {
            PopulateProfiles(savedStateFile);
            PopulateExtensions();
        }

        public void SortExtensions()
        {
            // Sort by file extension, but group by our handler, chained handlers, other handlers, and finally no handler
            // This uses a Sort extension to ObservableCollection
            extensions.Sort((e, f) =>
            {
                if (e.PropertyHandlerState != f.PropertyHandlerState)
                { 
                    if (e.PropertyHandlerState == HandlerState.Ours)
                        return -1;
                    else if (f.PropertyHandlerState == HandlerState.Ours)
                        return 1;
                    else if (e.PropertyHandlerState == HandlerState.Chained)
                        return -1;
                    else if (f.PropertyHandlerState == HandlerState.Chained)
                        return 1;
                    else if (e.PropertyHandlerState == HandlerState.Foreign)
                        return -1;
                    else if (f.PropertyHandlerState == HandlerState.Foreign)
                        return 1;
                }
                return e.Name.CompareTo(f.Name);
            });
            //Extensions.NotifyReset();
        }

#if CmdLine
#else
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
#endif

        public void LoadSavedState(string savedStateFile)
        {
            var fiDefault = GetDefaultSavedStateInfo();

            // If a state file has been specified, use it 
            if (savedStateFile != null)
            {
                var fi = new FileInfo(savedStateFile);
                if (!fi.Exists)
                    throw new AssocMgrException
                    {
                        Description = String.Format(LocalizedMessages.MissingDefinitionsFile, savedStateFile),
                        Exception = null,
                        ErrorCode = WindowsErrorCodes.ERROR_FILE_NOT_FOUND
                    };

                savedState = LoadSavedState(fi); 

                // If it's not just the default one, remember that we used a non-default file
                if (String.Compare(savedStateFile, fiDefault.FullName, true) != 0)
                    nonDefaultStateFileLoaded = true;
            }
            else if (fiDefault.Exists)
            {
                savedState = LoadSavedState(fiDefault); 
            }
        }

        public void StoreUpdatedProfile(Profile p)
        {
            if (!nonDefaultStateFileLoaded)
                StoreSavedState();
            else
            {
                SavedState newState;
                var fiDefault = GetDefaultSavedStateInfo();

                if (fiDefault.Exists)
                {
                    newState = LoadSavedState(fiDefault);
                    newState.CustomProfiles.RemoveAll(x => x.Name == p.Name);
                }
                else
                    newState = new SavedState();

                newState.CustomProfiles.Add(p);
                StoreSavedStateAsDefault(newState);
            }
        }

        public void StoreSavedState()
        {
            StoreSavedStateAsDefault(savedState);
        }

        private FileInfo GetDefaultSavedStateInfo()
        {
            DirectoryInfo di = ObtainDataDirectory();
            return new FileInfo(di.FullName + @"\SavedState.xml");
        }

        private SavedState LoadSavedState(FileInfo fi)
        {
            try
            {
                XmlSerializer x = new XmlSerializer(typeof(SavedState));
                using (TextReader reader = new StreamReader(fi.FullName))
                {
                    return (SavedState)x.Deserialize(reader);
                }
            }
            catch (Exception ex)
            {
                if (ex is AssocMgrException)
                    throw ex;
                else
                    throw new AssocMgrException { Description = LocalizedMessages.XmlParseError, Exception = ex, ErrorCode = WindowsErrorCodes.ERROR_XML_PARSE_ERROR };
            }
        }

        private void StoreSavedStateAsDefault(SavedState state)
        {
            var fi = GetDefaultSavedStateInfo();
            try
            {
                XmlSerializer x = new XmlSerializer(typeof(SavedState));
                using (TextWriter writer = new StreamWriter(fi.FullName))
                {
                    x.Serialize(writer, state);
                }
            }
            catch (Exception ex)
            {
                throw new AssocMgrException { Description = LocalizedMessages.XmlWriteError, Exception = ex, ErrorCode = WindowsErrorCodes.ERROR_XML_PARSE_ERROR };
            }
        }

        public Extension CreateExtension(string name)
        {
            Extension e = new Extension() { Name = name, State = this };
            e.RecordPropertyHandler(null, null);
            dictExtensions.Add(name.ToLower(), e);
            Extensions.Add(e);

            return e;
        }

        public Extension GetExtensionByName(string name)
        {
            Extension e;
            if (dictExtensions.TryGetValue(name.ToLower(), out e))
                return e;
            else
                return null;
        }

        public IEnumerable<Extension> GetExtensionsUsingProfile(Profile p)
        {
            return Extensions.Where(e => e.Profile == p);
        }

        public Profile GetProfileByName(string name)
        {
            return BuiltInProfiles.Concat(CustomProfiles).FirstOrDefault(x => String.Compare(x.Name, name, true) == 0);
        }

        private void PopulateProfiles(string savedStateFile)
        {
            foreach (Profile p in Profile.GetBuiltinProfiles(this))
                BuiltInProfiles.Add(p);

            LoadSavedState(savedStateFile);
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
            using (RegistryKey handlers = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\PropertySystem\PropertyHandlers", false))
            {
                foreach (string name in handlers.GetSubKeyNames())
                {
                    using (RegistryKey key = handlers.OpenSubKey(name, false))
                    {
                        string handlerGuid = (string)key.GetValue(null);
                        string handlerChainedGuid = (string)key.GetValue(Extension.ChainedValueName);

                        Extension e;
                        if (dictExtensions.TryGetValue(name.ToLower(), out e))
                        {
                            e.RecordPropertyHandler(handlerGuid, handlerChainedGuid);

                            e.IdentifyCurrentProfile();
                        }
                    }
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
