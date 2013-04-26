// Copyright (c) 2013, Dijii, and released under the Common Public License.  This, with other relevant licenses, can be found in the root of this distribution.

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;

namespace FileMetadataAssociationManager
{
    public class ObservableCollectionWithReset<T> : ObservableCollection<T>
    {
        public void NotifyReset()
        {
            OnCollectionChanged(new NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Reset));
        }
    }
    
}
