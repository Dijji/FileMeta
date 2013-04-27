// Copyright (c) 2013, Dijji, and released under Ms-PL.  This, with other relevant licenses, can be found in the root of this distribution.

using System;
using System.Collections.Generic;
using System.ComponentModel;


namespace FileMetadataAssociationManager
{
    public class TreeItem : INotifyPropertyChanged
    {
        string name = null;

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged(String info)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(info));
            }
        }

        public TreeItem(string name)
        {
            this.name = name;
        }

        List<TreeItem> children = new List<TreeItem>();

        public string Name { get { return name; } }
        public List<TreeItem> Children { get { return children; } }
    }
}
