// Copyright (c) 2013, Dijji, and released under Ms-PL.  This, with other relevant licenses, can be found in the root of this distribution.

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Xml.Serialization;
using FileMetadataAssociationManager.Resources;

namespace FileMetadataAssociationManager
{
    public class Profile
    {
        const string FullDetailsOfficeProfile = "prop:System.PropGroup.Description;System.Title;System.Subject;System.Keywords;System.Category;System.Comment;System.Rating;System.PropGroup.Origin;System.Author;System.Document.LastAuthor;System.Document.RevisionNumber;System.Document.Version;System.ApplicationName;System.Company;System.Document.Manager;System.Document.DateCreated;System.Document.DateSaved;System.Document.DatePrinted;System.Document.TotalEditingTime;System.PropGroup.Content;System.ContentStatus;System.ContentType;System.Document.PageCount;System.Document.WordCount;System.Document.CharacterCount;System.Document.LineCount;System.Document.ParagraphCount;System.Document.Template;System.Document.Scale;System.Document.LinksDirty;System.Language;System.PropGroup.FileSystem;System.ItemNameDisplay;System.ItemType;System.ItemFolderPathDisplay;System.DateCreated;System.DateModified;System.Size;System.FileAttributes;System.OfflineAvailability;System.OfflineStatus;System.SharedWith;System.FileOwner;System.ComputerName";
        const string PreviewDetailsOfficeProfile = "prop:*System.DateModified;System.Author;System.Keywords;System.Rating;*System.Size;System.Title;System.Comment;System.Category;*System.Document.PageCount;System.ContentStatus;System.ContentType;*System.OfflineAvailability;*System.OfflineStatus;System.Subject;*System.DateCreated;*System.SharedWith";
        const string InfoTipOfficeProfile = "prop:System.ItemTypeText;System.Size;System.DateModified;System.Document.PageCount";

        const string FullDetailsSimpleProfile = "prop:System.PropGroup.Description;System.Title;System.Subject;System.Keywords;System.Category;System.Comment;System.Rating;System.PropGroup.Origin;System.Author;System.Document.RevisionNumber";
        const string PreviewDetailsSimpleProfile = "prop:System.Title;System.Subject;System.Keywords;System.Category;System.Comment;System.Rating;System.Author;System.Document.RevisionNumber";
        const string InfoTipSimpleProfile = "prop:System.ItemTypeText;System.Size;System.DateModified;System.Comment";

        // These three properties are the 'master' representation of the details for the profile
        private ObservableCollection<TreeItem> fullDetails = new ObservableCollection<TreeItem>();
        private ObservableCollection<PropertyListEntry> previewDetails = new ObservableCollection<PropertyListEntry>();
        private ObservableCollection<PropertyListEntry> infoTips = new ObservableCollection<PropertyListEntry>();
        private bool isReadOnly = false;

        public string Name { get; set; }

        // These three properties get and set the registry and XML form of the details for the profile
        public string FullDetailsString { get { return GetFullDetailsString(); } set { ParseFullDetailsString(value); } }
        public string PreviewDetailsString { get { return GetPropertiesString(PreviewDetails); } set { ParsePropertiesString(value, PreviewDetails); } }
        public string InfoTipString { get { return GetPropertiesString(InfoTips); } set { ParsePropertiesString(value, InfoTips); } }

        [XmlIgnore]
        public State State { get; set; }

        [XmlIgnore]
        public Profile Original { get; set; }

        [XmlIgnore]
        public ObservableCollection<TreeItem> FullDetails { get { return fullDetails; } }

        [XmlIgnore]
        public ObservableCollection<PropertyListEntry> PreviewDetails { get { return previewDetails; } }

        [XmlIgnore]
        public ObservableCollection<PropertyListEntry> InfoTips { get { return infoTips; } }

        [XmlIgnore]
        public bool IsReadOnly { get { return isReadOnly; } }

        [XmlIgnore]
        public bool IsNull { get; set; }

        public Profile()
        {
        }

        // Create a clone of this profile
        public Profile CreateClone()
        {
            Profile p = new Profile {Name = this.Name, State = this.State, FullDetailsString = this.FullDetailsString,
                                     PreviewDetailsString = this.PreviewDetailsString, InfoTipString = this.InfoTipString };
            return p;
        }

        // Does this profile differ From another one?
        public bool DiffersFrom(Profile p)
        {
            return this.Name != p.Name ||
                   this.FullDetailsString != p.FullDetailsString ||
                   this.PreviewDetailsString != p.PreviewDetailsString ||
                   this.InfoTipString != p.InfoTipString;
        }

        // Copy the values from another profile
        public void UpdateFrom(Profile p)
        {
            this.Name = p.Name;
            this.FullDetailsString = p.FullDetailsString;
            this.PreviewDetailsString = p.PreviewDetailsString;
            this.InfoTipString = p.InfoTipString;
        }

        // Merge the values from another profile
        public void MergeFrom(Profile p)
        {
            MergeFullDetails(p.FullDetails);
            MergeProperties(p.PreviewDetails, PreviewDetails);
            MergeProperties(p.InfoTips, InfoTips);
        }

        public bool HasGroupInFullDetails(string group)
        {
            // Walk the top level of the tree
            return (FullDetails.FirstOrDefault(ti => ti.Name == group) != null);
        }

        public bool HasPropertyInFullDetails(string property)
        {
            // Walk the the tree
            foreach (TreeItem ti in FullDetails)
            {
                if (ti.Children.FirstOrDefault(tic =>  ModuloAsterisk(tic.Name) == property) != null)
                    return true;
            }

            return false;
        }

        public bool HasPropertyInPreviewDetails(string property)
        {
            // Check properties modulo the presence of an '*' prefix
            return HasPropertyInProperties(property, PreviewDetails);
        }

        public bool HasPropertyInInfoTip(string property)
        {
            // Check properties modulo the presence of an '*' prefix
            return HasPropertyInProperties(property, InfoTips);
        }

        public void AddFullDetailsGroup(string group, TreeItem target, bool before)
        {
            TreeItem toAdd = new TreeItem(group, PropType.Group);
            InsertGroup(toAdd, target, before);
        }

        public void MoveFullDetailsGroup(TreeItem toMove, TreeItem target, bool before)
        {
            if (toMove != target)
            {
                fullDetails.Remove(toMove);
                InsertGroup(toMove, target, before);
            }
        }

        public void AddFullDetailsProperty(string property, TreeItem target, bool before)
        {
            TreeItem toAdd = new TreeItem(property, PropType.Normal);
            InsertFullDetailsProperty(toAdd, target, before);
        }

        public void MoveFullDetailsProperty(TreeItem toMove, TreeItem target, bool before)
        {
            if (toMove != target)
            {
                toMove.Parent.RemoveChild(toMove);
                InsertFullDetailsProperty(toMove, target, before);
            }
        }

        public void AddPreviewDetailsProperty(string property, PropertyListEntry target, bool before)
        {
            InsertPropertyInProperties(PreviewDetails, new PropertyListEntry(property), target, before);
        }

        public void MovePreviewDetailsProperty(PropertyListEntry toMove, PropertyListEntry target, bool before)
        {
            if (toMove != target)
            {
                PreviewDetails.Remove(toMove);
                InsertPropertyInProperties(PreviewDetails, toMove, target, before);
            }
        }

        public void AddInfoTipProperty(string property, PropertyListEntry target, bool before)
        {
            InsertPropertyInProperties(InfoTips, new PropertyListEntry(property), target, before);
        }

        public void MoveInfoTipProperty(PropertyListEntry toMove, PropertyListEntry target, bool before)
        {
            if (toMove != target)
            {
                InfoTips.Remove(toMove);
                InsertPropertyInProperties(InfoTips, toMove, target, before);
            }
        }

        private void InsertGroup(TreeItem toInsert, TreeItem target, bool before)
        {
            if (target == null)
                FullDetails.Add(toInsert);
            else
            {
                int index;

                if ((PropType)target.Item == PropType.Group)
                    index = FullDetails.IndexOf(target) + (before ? 0 : 1);
                else
                    index = FullDetails.IndexOf(target.Parent) + 1;

                FullDetails.Insert(index, toInsert);
            }
        }

        private void InsertFullDetailsProperty(TreeItem toInsert, TreeItem target, bool before)
        {
            if (target == null)
                // Drop in the last group
                FullDetails.Last().AddChild(toInsert);

            else if ((PropType)target.Item == PropType.Group)
            {
                if (before)
                {
                    int index = FullDetails.IndexOf(target);
                    if (index > 0)
                        index--;

                    FullDetails[index].AddChild(toInsert);
                }
                else
                    target.InsertChild(0, toInsert);
            }
            else
            {
                int index = target.Parent.Children.IndexOf(target) + (before ? 0 : 1);
                if (index < target.Parent.Children.Count)
                    target.Parent.InsertChild(index, toInsert);
                else
                    target.Parent.AddChild(toInsert);
            }
        }

        public void RemoveFullDetailsItem(TreeItem toRemove)
        {
            if ((PropType)toRemove.Item == PropType.Group)
            {
                foreach (var ti in toRemove.Children)
                {
                    RemovePreviewDetailsProperty(ti.Name);
                    RemoveInfoTipProperty(ti.Name);
                }

                FullDetails.Remove(toRemove);
            }
            else
            {
                RemovePreviewDetailsProperty(toRemove.Name);
                RemoveInfoTipProperty(toRemove.Name);
                toRemove.Parent.RemoveChild(toRemove);
            }
        }

        public void ToggleAsteriskFullDetailsItem(TreeItem toToggle)
        {
            if ((PropType)toToggle.Item == PropType.Normal)
            {
                if (toToggle.Name.StartsWith("*"))
                    toToggle.ChangeName(toToggle.Name.Substring(1));
                 else
                    toToggle.ChangeName("*" + toToggle.Name);
            }
        }

        public void RemovePreviewDetailsProperty(string property)
        {
            RemovePropertyFromProperties(property, PreviewDetails);
        }

        public void RemovePreviewDetailsProperty(PropertyListEntry property)
        {
            PreviewDetails.Remove(property);
        }

        public void RemoveInfoTipProperty(string property)
        {
            RemovePropertyFromProperties(property, InfoTips);
        }

        public void RemoveInfoTipProperty(PropertyListEntry property)
        {
            InfoTips.Remove(property);
        }

        private void InsertPropertyInProperties(ObservableCollection<PropertyListEntry> properties, PropertyListEntry toInsert, PropertyListEntry target, bool before)
        {
            if (target == null)
                // Drop at the end
                properties.Add(toInsert);
            else
            {
                int index = properties.IndexOf(target) + (before ? 0 : 1);
                if (index < properties.Count)
                    properties.Insert(index, toInsert);
                else
                    properties.Add(toInsert);
            }
        }
        
        private void ParseFullDetailsString(string registryEntry)
        {
            FullDetails.Clear();

            Dictionary<string, TreeItem> groups = new Dictionary<string, TreeItem>();

             if (registryEntry.StartsWith("prop:"))
                registryEntry = registryEntry.Remove(0, 5);
            
            string[] props = registryEntry.Split(';');
            TreeItem curr = null;

            foreach (string prop in props)
            {
                string[] parts = prop.Split('.');

                if (parts.Length == 3 && parts[0] == "System" && parts[1] == "PropGroup")
                {
                    // put following entries under the specified group
                    string group = parts[2];
                    if (!groups.TryGetValue(group, out curr))
                    {
                        // Mark the TreeItem as a group property
                        curr = new TreeItem(group, PropType.Group);
                        groups.Add(group, curr);
                    }
                }
                else if (curr != null)
                {
                    // Mark the TreeItem as a normal property
                    curr.AddChild(new TreeItem(prop, PropType.Normal));
                }
            }

            // Add the tree item roots to the displayable collection
            foreach (var ti in groups.Values)
                FullDetails.Add(ti);
        }

        private void MergeFullDetails(IEnumerable<TreeItem> source)
        {
            foreach (var tiSource in source)
            {
                if (((PropType)tiSource.Item) == PropType.Group)
                {
                    var tiTarget = FullDetails.Where(ti => ti.Name == tiSource.Name).FirstOrDefault();
                    if (tiTarget == null)
                        FullDetails.Add(tiSource.Clone());
                    else
                    {
                        foreach (var ti in tiSource.Children)
                        {
                            string s = ModuloAsterisk(ti.Name);
                            if (tiTarget.Children.FirstOrDefault(t => ModuloAsterisk(t.Name) == s) == null)
                            {
                                tiTarget.AddChild(ti.Clone());
                            }
                        }
                    }
                }
            }
        }

        private string GetFullDetailsString()
        {
            StringBuilder sb = new StringBuilder("prop:");
            bool first = true;

            foreach (var ti in FullDetails)
            {
                if (!first)
                    sb.Append(";");
                else
                    first = false;

                if ((PropType)ti.Item == PropType.Group)
                {
                    sb.Append("System.PropGroup.");
                    sb.Append(ti.Name);

                    foreach (var tic in ti.Children)
                    {
                        if (!first)
                            sb.Append(";");
                        sb.Append(tic.Name);
                    }
                }
                else
                    sb.Append(ti.Name);
            }

            return sb.ToString();
        }

        private void ParsePropertiesString(string registryEntry, ObservableCollection<PropertyListEntry> properties)
        {
            properties.Clear();

            if (registryEntry.StartsWith("prop:"))
                registryEntry = registryEntry.Remove(0, 5);

            string[] props = registryEntry.Split(';');

            // This a simple list pf semicolon delimited names
            // Some property names are prefixed with '*', we preserve it, but we don't do it ourselves
            // MSDN says it means 'Do not show the property in the Preview Pane as instructed in the PreviewDetails registry key value',
            // but it never seems to have any effect
            foreach (string prop in props)
            {
                if (prop.Length == 0)
                    continue;
                else 
                    properties.Add(new PropertyListEntry(prop));
            }
        }

        public bool HasPropertyInProperties(string property, IEnumerable<PropertyListEntry> properties)
        {
            // Check properties modulo the presence of an '*' prefix
            var s = ModuloAsterisk(property);
            return (properties.FirstOrDefault(pe => pe.Name == s) != null);
        }

        private void MergeProperties(IEnumerable<PropertyListEntry> source, ObservableCollection<PropertyListEntry> target)
        {
            var result = target.Union<PropertyListEntry>(source, (a,b) => a.Name == b.Name).ToList();

            target.Clear();
            foreach (var s in result)
                target.Add(s);
        }

        private void RemovePropertyFromProperties(string property, ObservableCollection<PropertyListEntry> properties)
        {
            string s = ModuloAsterisk(property);
            int index = properties.FindIndex(pe => pe.Name == s);
            if (index >= 0)
                properties.RemoveAt(index);
        }

        private string GetPropertiesString(IEnumerable<PropertyListEntry> properties)
        {
            StringBuilder sb = new StringBuilder("prop:");
            bool first = true;

            foreach (PropertyListEntry prop in properties)
            {
                if (!first)
                    sb.Append(";");
                else
                    first = false;
                sb.Append(prop.NameString);
            }

            return sb.ToString();
        }

        private string ModuloAsterisk(string s)
        {
            if (s.StartsWith("*"))
                return s.Substring(1);
            else
                return s;
        }

        public static List<Profile> GetBuiltinProfiles(State state)
        {
            List<Profile> ps = new List<Profile>();

            Profile p = new Profile { Name = "Simple", State = state, FullDetailsString = FullDetailsSimpleProfile,
                                      PreviewDetailsString = PreviewDetailsSimpleProfile, InfoTipString = InfoTipSimpleProfile };
            p.isReadOnly = true;
            ps.Add(p);

            p = new Profile { Name = "Office DSOfile", State = state, FullDetailsString = FullDetailsOfficeProfile,
                              PreviewDetailsString = PreviewDetailsOfficeProfile, InfoTipString = InfoTipOfficeProfile };
            p.isReadOnly = true;
            ps.Add(p);

            return ps;
        }

        // Construct the nth new name
        public static string NewName(int index)
        {
            if (index == 1)
                return LocalizedMessages.NewProfileName;
            else 
                return String.Format("{0} ({1})", LocalizedMessages.NewProfileName, index);
        }


    }
}
