// Copyright (c) 2013, Dijii, and released under the Common Public License.  This, with other relevant licenses, can be found in the root of this distribution.

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;

namespace FileMetadataAssociationManager
{
    public class Profile
    {
        const string FullDetailsOfficeProfile = "prop:System.PropGroup.Description;System.Title;System.Subject;System.Keywords;System.Category;System.Comment;System.Rating;System.PropGroup.Origin;System.Author;System.Document.LastAuthor;System.Document.RevisionNumber;System.Document.Version;System.ApplicationName;System.Company;System.Document.Manager;System.Document.DateCreated;System.Document.DateSaved;System.Document.DatePrinted;System.Document.TotalEditingTime;System.PropGroup.Content;System.ContentStatus;System.ContentType;System.Document.PageCount;System.Document.WordCount;System.Document.CharacterCount;System.Document.LineCount;System.Document.ParagraphCount;System.Document.Template;System.Document.Scale;System.Document.LinksDirty;System.Language;System.PropGroup.FileSystem;System.ItemNameDisplay;System.ItemType;System.ItemFolderPathDisplay;System.DateCreated;System.DateModified;System.Size;System.FileAttributes;System.OfflineAvailability;System.OfflineStatus;System.SharedWith;System.FileOwner;System.ComputerName";
        const string PreviewDetailsOfficeProfile = "prop:*System.DateModified;System.Author;System.Keywords;System.Rating;*System.Size;System.Title;System.Comment;System.Category;*System.Document.PageCount;System.ContentStatus;System.ContentType;*System.OfflineAvailability;*System.OfflineStatus;System.Subject;*System.DateCreated;*System.SharedWith";

        const string FullDetailsSimpleProfile = "prop:System.PropGroup.Description;System.Title;System.Subject;System.Keywords;System.Category;System.Comment;System.Rating;System.PropGroup.Origin;System.Author;System.Document.RevisionNumber";
        const string PreviewDetailsSimpleProfile = "prop:System.Title;System.Subject;System.Keywords;System.Category;System.Comment;System.Rating;System.Author;System.Document.RevisionNumber";

        private ObservableCollection<TreeItem> fullDetails = new ObservableCollection<TreeItem>();
        private ObservableCollection<string> previewDetails = new ObservableCollection<string>();
        private string fullDetailsString;
        private string previewDetailsString;

        public string Name { get; set; }
        public State State { get; set; }
        public ObservableCollection<TreeItem> FullDetails { get { return fullDetails; } }
        public ObservableCollection<string> PreviewDetails { get { return previewDetails; } }
        public string FullDetailsString { get { return fullDetailsString; } }
        public string PreviewDetailsString { get { return previewDetailsString; } }

        public void ParseFullDetails(string registryEntry)
        {
            fullDetailsString = registryEntry;

            Dictionary<string, TreeItem> groups = new Dictionary<string, TreeItem>();

            // Seed the tree with the known groups
            foreach (string g in State.GroupProperties)
            {
                groups.Add(g, new TreeItem(g));
            }

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
                    if (groups.ContainsKey(parts[2]))
                        curr = groups[parts[2]];
                }
                else if (curr != null)
                {
                    curr.Children.Add(new TreeItem(prop));
                }
            }

            // Add the tree item roots to the displayable collection
            foreach (var ti in groups.Values)
                FullDetails.Add(ti);
        }

        public void ParsePreviewDetails(string registryEntry)
        {
            previewDetailsString = registryEntry;

            if (registryEntry.StartsWith("prop:"))
                registryEntry = registryEntry.Remove(0, 5);

            string[] props = registryEntry.Split(';');

            // This a simple list pf semicolon deleimited names
            // Some DSOFile names are prefixed with '*', but I don't know what this means
            foreach (string prop in props)
            {
                PreviewDetails.Add(prop);
            }
        }

        public static List<Profile> GetBuiltinProfiles(State state)
        {
            List<Profile> ps = new List<Profile>();
            Profile p = new Profile { Name = "Office DSOfile" };
            p.State = state;
            p.ParseFullDetails(FullDetailsOfficeProfile);
            p.ParsePreviewDetails(PreviewDetailsOfficeProfile);
            ps.Add(p);

            p = new Profile { Name = "Simple" };
            p.State = state;
            p.ParseFullDetails(FullDetailsSimpleProfile);
            p.ParsePreviewDetails(PreviewDetailsSimpleProfile);
            ps.Add(p);

            return ps;
        }
    }
}
