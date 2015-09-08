// Copyright (c) 2015, Dijji, and released under Ms-PL.  This, with other relevant licenses, can be found in the root of this distribution.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FileMetadataAssociationManager
{
    public class SavedState
    {
        private List<Profile> customProfiles = new List<Profile>();

        public List<Profile> CustomProfiles { get { return customProfiles; } }

    }
}
