using System;

namespace FileMetadataAssociationManager
{
    public class NameChangedEventArgs : EventArgs
    {
        private string newName;

        public NameChangedEventArgs(string newName)
        {
            this.newName = newName;
        }

        public virtual string NewName { get { return newName; } }
    }

    public delegate void NameChangedEventHandler(object sender, NameChangedEventArgs e);
}
