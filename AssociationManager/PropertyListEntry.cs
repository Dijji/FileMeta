// Copyright (c) 2016, Dijji, and released under Ms-PL.  This, with other relevant licenses, can be found in the root of this distribution.

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace FileMetadataAssociationManager
{
    public class PropertyListEntry : INotifyPropertyChanged
    {
        public bool Asterisk { get; set; }
        public string Name { get; set; }
        public string NameString { get { return Asterisk ? "*" + Name : Name; } }

        public PropertyListEntry(string nameString)
        {
            if (nameString.StartsWith("*"))
            {
                Asterisk = true;
                Name = nameString.Substring(1);
            }
            else
            {
                Asterisk = false;
                Name = nameString;
            }
        }

        public void ToggleAsterisk()
        {
            Asterisk = !Asterisk;
            OnPropertyChanged("NameString");
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
