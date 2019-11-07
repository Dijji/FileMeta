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
using System.Diagnostics;

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
            List<SystemProperty> systemProperties = new List<SystemProperty>();
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

                    for (uint i = 0; i < count; i++)
                    {
                        propertyDescriptionList.GetAt(i, ref guid, out propertyDescription);

                        string propName;
                        propertyDescription.GetCanonicalName(out propName);
                        IntPtr displayNamePtr;
                        string displayName = null;
                        propertyDescription.GetDisplayName(out displayNamePtr);
                        if (displayNamePtr != IntPtr.Zero)
                            displayName = Marshal.PtrToStringAuto(displayNamePtr);
                        SystemProperty sp = new SystemProperty {FullName = propName, DisplayName = displayName };

                        systemProperties.Add(sp);

                        // If we ever need it:
                        // ShellPropertyDescription desc = new ShellPropertyDescription(propertyDescription);

                        if (propertyDescription != null)
                        {
                            Marshal.ReleaseComObject(propertyDescription);
                            propertyDescription = null;
                        }
                    }

                    Dictionary<string, TreeItem> dict = new Dictionary<string, TreeItem>();
                    List<TreeItem> roots = new List<TreeItem>();

                    // Build tree based on property names
                    foreach (SystemProperty sp in systemProperties)
                    {
                        AddTreeItem(dict, roots, sp);
                    }

                    // Wire trees to tree controls, tweaking the structure as we go
                    TreeItem propGroup = null;
                    foreach (TreeItem root in roots)
                    {
                        if (root.Name == "System")
                        {
                            AllProperties.Insert(0, root);

                            // Move property groups from root to their own list
                            propGroup = root.Children.Where(x => x.Name == "PropGroup").FirstOrDefault();
                            if (propGroup != null)
                            {
                                foreach (TreeItem ti in propGroup.Children)
                                    GroupProperties.Add(ti.Name);
                                root.RemoveChild(propGroup);
                            }

                            // Make remaining children of System that are parents into roots
                            List<TreeItem> movers = new List<TreeItem>(root.Children.Where(x => x.Children.Count() > 0));
                            foreach (TreeItem ti in movers)
                            {
                                root.RemoveChild(ti);
                                AllProperties.Add(ti);
                            }
                        }
                        else
                            AllProperties.Add(root);
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

        // Top level entry point for the algorithm that builds the property name tree from an unordered sequence
        // of property names
        private TreeItem AddTreeItem(Dictionary<string, TreeItem> dict, List<TreeItem> roots, SystemProperty sp)
         {
            Debug.Assert(sp.FullName.Contains('.')); // Because the algorithm assumes that this is the case
            TreeItem ti = AddTreeItemInner(dict, roots, sp.FullName, sp.DisplayName);
            ti.Item = sp;

            return ti;
        }

        // Recurse backwards through each term in the property name, adding tree items as we go,
        // until we join onto an existing part of the tree
        private TreeItem AddTreeItemInner(Dictionary<string, TreeItem> dict, List<TreeItem> roots, 
            string name, string displayName = null)
        {
            TreeItem ti, parent;
            string parentName = FirstPartsOf(name);

            if (parentName != null)
            {
                if (!dict.TryGetValue(parentName, out parent))
                {
                    parent = AddTreeItemInner(dict, roots, parentName);
                    dict.Add(parentName, parent);
                }

                if (displayName != null)
                    ti = new TreeItem(String.Format("{0} ({1})", LastPartOf(name), displayName));
                else
                    ti = new TreeItem(LastPartOf(name));

                parent.AddChild(ti);
            }
            else
            {
                if (!dict.TryGetValue(name, out ti))
                {
                    ti = new TreeItem(name);
                    roots.Add(ti);
                }
            }

            return ti;
        }

        private string FirstPartsOf(string name)
        {
            int index = name.LastIndexOf('.');
            return index >= 0 ? name.Substring(0, index) : null;
        }

        private string LastPartOf(string name)
        {
            int index = name.LastIndexOf('.');
            return index >= 0 ? name.Substring(index + 1) : name;
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

    public class SystemProperty
    {
        public string FullName { get; set; }
        public string DisplayName { get; set; }
    }
}
