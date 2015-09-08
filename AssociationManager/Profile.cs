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

        const string FullDetailsSimpleProfile = "prop:System.PropGroup.Description;System.Title;System.Subject;System.Keywords;System.Category;System.Comment;System.Rating;System.PropGroup.Origin;System.Author;System.Document.RevisionNumber";
        const string PreviewDetailsSimpleProfile = "prop:System.Title;System.Subject;System.Keywords;System.Category;System.Comment;System.Rating;System.Author;System.Document.RevisionNumber";

        // These two properties are the 'master' representation of the details for the profile
        private ObservableCollection<TreeItem> fullDetails = new ObservableCollection<TreeItem>();
        private ObservableCollection<string> previewDetails = new ObservableCollection<string>();
        private bool isReadOnly = false;

        public string Name { get; set; }

        // These two properties get and set the registry and XML form of the details for the profile
        public string FullDetailsString { get { return GetFullDetailsString(); } set { ParseFullDetailsString(value); } }
        public string PreviewDetailsString { get { return GetPreviewDetailsString(); } set { ParsePreviewDetailsString(value); } }

        [XmlIgnore]
        public State State { get; set; }

        [XmlIgnore]
        public Profile Original { get; set; }

        [XmlIgnore]
        public ObservableCollection<TreeItem> FullDetails { get { return fullDetails; } }

        [XmlIgnore]
        public ObservableCollection<string> PreviewDetails { get { return previewDetails; } }

        [XmlIgnore]
        public bool IsReadOnly { get { return isReadOnly; } }

        public Profile()
        {
        }

        // Create a clone of this profile
        public Profile CreateClone()
        {
            Profile p = new Profile {Name = this.Name, State = this.State, 
                                     FullDetailsString = this.FullDetailsString, PreviewDetailsString = this.PreviewDetailsString };
            return p;
        }

        // Does this profile differ From another one?
        public bool DiffersFrom(Profile p)
        {
            return this.Name != p.Name ||
                   this.FullDetailsString != p.FullDetailsString ||
                   this.PreviewDetailsString != p.PreviewDetailsString;
        }

        // Copy the values from another profile
        public void UpdateFrom(Profile p)
        {
            this.Name = p.Name;
            this.FullDetailsString = p.FullDetailsString;
            this.PreviewDetailsString = p.PreviewDetailsString;
        }

        public bool HasGroupInFullDetails(string group)
        {
            // Walk the top level of the tree
            foreach (TreeItem ti in FullDetails)
            {
                if (ti.Name == group)
                    return true;
            }
            return false;
        }

        public bool HasPropertyInFullDetails(string property)
        {
            // Walk the the tree
            foreach (TreeItem ti in FullDetails)
            {
                foreach (TreeItem tic in ti.Children)
                {
                    if (tic.Name == property)
                        return true;
                }
            }

            return false;
        }

        public bool HasPropertyInPreviewDetails(string property)
        {
            // Check properties modulo the presence of an '*' prefix
            return PreviewDetails.Contains(property) ||
                   PreviewDetails.Contains("*" + property);
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

        public void AddPreviewDetailsProperty(string property, string target, bool before)
        {
            InsertPreviewDetailsProperty(property, target, before);
        }

        public void MovePreviewDetailsProperty(string toMove, string target, bool before)
        {
            if (toMove != target)
            {
                PreviewDetails.Remove(toMove);
                InsertPreviewDetailsProperty(toMove, target, before);
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
                    RemovePreviewDetailsProperty(ti.Name);

                FullDetails.Remove(toRemove);
            }
            else
            {
                RemovePreviewDetailsProperty(toRemove.Name);
                toRemove.Parent.RemoveChild(toRemove);
            }
        }

        public void RemovePreviewDetailsProperty(string property)
        {
            int index = PreviewDetails.IndexOf(property);
            if (index == -1 )
                index = PreviewDetails.IndexOf("*" + property);
            if (index >= 0)
                PreviewDetails.RemoveAt(index);
        }

        private void InsertPreviewDetailsProperty(string toInsert, string target, bool before)
        {
            if (target == null)
                // Drop at the end
                PreviewDetails.Add(toInsert);
            else
            {
                int index = PreviewDetails.IndexOf(target) + (before ? 0 : 1);
                if (index < PreviewDetails.Count)
                    PreviewDetails.Insert(index, toInsert);
                else
                    PreviewDetails.Add(toInsert);
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

        private void ParsePreviewDetailsString(string registryEntry)
        {
            PreviewDetails.Clear();

            if (registryEntry.StartsWith("prop:"))
                registryEntry = registryEntry.Remove(0, 5);

            string[] props = registryEntry.Split(';');

            // This a simple list pf semicolon delimited names
            // Some DSOFile names are prefixed with '*', but I don't know what this means, and we don't do it ourselves
            foreach (string prop in props)
            {
                PreviewDetails.Add(prop);
            }
        }

        private string GetPreviewDetailsString()
        {
            StringBuilder sb = new StringBuilder("prop:");
            bool first = true;

            foreach (string prop in PreviewDetails)
            {
                if (!first)
                    sb.Append(";");
                else
                    first = false;
                sb.Append(prop);
            }

            return sb.ToString();
        }

        public static List<Profile> GetBuiltinProfiles(State state)
        {
            List<Profile> ps = new List<Profile>();

            Profile p = new Profile { Name = "Simple", State = state, 
                                      FullDetailsString = FullDetailsSimpleProfile, PreviewDetailsString =  PreviewDetailsSimpleProfile};
            p.isReadOnly = true;
            ps.Add(p);

            p = new Profile { Name = "Office DSOfile", State = state, 
                              FullDetailsString = FullDetailsOfficeProfile, PreviewDetailsString = PreviewDetailsOfficeProfile };
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
