// Copyright (c) 2013, Dijji, and released under Ms-PL.  This, with other relevant licenses, can be found in the root of this distribution.

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;


namespace FileMetadataAssociationManager
{
    public class TreeItem : INotifyPropertyChanged
    {
        string name = null;
        bool isSelected = false;
        TreeItem parent = null;
        ObservableCollection<TreeItem> children = new ObservableCollection<TreeItem>();

        public event PropertyChangedEventHandler PropertyChanged;
        public event NameChangedEventHandler NameChanged;

        protected void OnPropertyChanged(String info)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(info));
            }
        }

        public TreeItem(string name, object item = null)
        {
            this.name = name;
            this.Item = item;
        }

        public string Name { get { return name; } }
        public object Item { get; set; }
        public TreeItem Parent { get { return parent; } }
        public ObservableCollection<TreeItem> Children { get { return children; } }

        public bool IsSelected
        { 
            get 
            {
                return isSelected;
            }
            set
            {
                if (value != isSelected)
                {
                    isSelected = value;
                    OnPropertyChanged("IsSelected");
                }
            }
        }

        public string EditableName
        {
            get 
            { 
                return Name; 
            }
            set
            {
                if (NameChanged != null)
                {
                    NameChanged(this, new NameChangedEventArgs(value));
                }
            }
        }

        public void AddChild(TreeItem child)
        {
            child.parent = this;
            Children.Add(child);
        }

        public void InsertChild(int index, TreeItem child)
        {
            child.parent = this;
            Children.Insert(index, child);
        }

        public void RemoveChild(TreeItem child)
        {
            child.parent = null;
            Children.Remove(child);
        }

        public void AbandonNameChange()
        {
            OnPropertyChanged("EditableName");
        }

        public void ChangeName(string newName)
        {
            name = newName;
            OnPropertyChanged("Name");
            OnPropertyChanged("EditableName");
        }

        public TreeItem Clone()
        {
            TreeItem clone = new TreeItem(this.Name, this.Item);
            foreach (var ti in this.Children)
                clone.AddChild(ti.Clone());
            return clone;
        }
    }
}
