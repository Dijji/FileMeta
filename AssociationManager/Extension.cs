// Copyright (c) 2013, Dijii, and released under Ms-PL.  This, with other relevant licenses, can be found in the root of this distribution.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
using System.Windows;
using System.Windows.Media;
using Microsoft.Win32;
using System.Security.Principal;
using FileMetadataAssociationManager.Resources;

namespace FileMetadataAssociationManager
{
    public enum HandlerState
    {
        None,
        Ours,
        Foreign,
        Chained,
    }

    public class Extension : INotifyPropertyChanged
    {
        const string OurPropertyHandlerTitle = "File Meta Property Handler";
        const string OurPropertyHandlerPrefix = "File Meta + ";
        const string OurPropertyHandlerGuid64 = "{D06391EE-2FEB-419B-9667-AD160D0849F3}";
        const string OurPropertyHandlerGuid32 = "{60211757-EF87-465e-B6C1-B37CF98295F9}";
        const string OurContextHandlerGuid64 = "{28D14D00-2D80-4956-9657-9D50C8BB47A5}";
        const string OurContextHandlerGuid32 = "{DA38301B-BE91-4397-B2C8-E27A0BD80CC5}";

        const string FullDetailsValueName = "FullDetails";
        const string PreviewDetailsValueName = "PreviewDetails";
        const string InfoTipValueName = "InfoTip";
        const string OldFullDetailsValueName = "FileMetaOldFullDetails";
        const string OldPreviewDetailsValueName = "FileMetaOldPreviewDetails";
        const string OldInfoTipValueName = "FileMetaOldInfoTip";
        const string FileMetaCustomProfileValueName = "FileMetaCustomProfile";
        const string ShellExKeyName = "ShellEx";
        const string ContextMenuHandlersKeyName = "ContextMenuHandlers";
        const string ContextHandlerKeyName = "FileMetadata";
        public const string ChainedValueName = "Chained";

        private static Nullable<bool> isOurPropertyHandlerRegistered = null;
        private static Nullable<bool> isOurContextHandlerRegistered = null;

        private Profile profile = null;  // The profile currently applied to this extension, if any

        private string PropertyHandlerGuid { get; set; } // hold as string because Guid.Parse() needs .Net 4
        private string PropertyHandlerTitle { get; set; }

        public string Name { get; set; }
        public State State { get; set; }
        public HandlerState PropertyHandlerState { get; private set; }
        public Profile Profile { get { return profile; } set { profile = value; } }
        public FontWeight Weight 
        {
            get 
            {
                switch (PropertyHandlerState)
                {
                    case HandlerState.Ours:
                    case HandlerState.Chained:
                        return FontWeights.ExtraBold;

                    case HandlerState.Foreign:
                    case HandlerState.None:
                    default:
                        return FontWeights.Normal;
                }
            } 
        }

        public string PropertyHandlerDisplay 
        {
            get 
            {
                switch (PropertyHandlerState)
                {
                    case HandlerState.Ours:
                        return OurPropertyHandlerTitle;

                    case HandlerState.Foreign:
                        return PropertyHandlerTitle != null ? PropertyHandlerTitle : PropertyHandlerGuid;

                    case HandlerState.Chained:
                        return OurPropertyHandlerPrefix + (PropertyHandlerTitle != null ? PropertyHandlerTitle : PropertyHandlerGuid);

                    case HandlerState.None:
                    default:
                        return LocalizedMessages.PropertyHandlerNone;
                }
            } 
        }

        public static bool IsElevated
        {
            get
            {
                return new WindowsPrincipal
                    (WindowsIdentity.GetCurrent()).IsInRole
                    (WindowsBuiltInRole.Administrator);
            }
        }

#if x64
        private static string OurPropertyHandlerGuid { get { return OurPropertyHandlerGuid64; } }
        private static string OurContextHandlerGuid { get { return OurContextHandlerGuid64; } }
#elif x86
        private static string OurPropertyHandlerGuid { get { return OurPropertyHandlerGuid32; } }
        private static string OurContextHandlerGuid { get { return OurContextHandlerGuid32; } }
#endif

        public static bool IsOurPropertyHandlerRegistered
        {
            get
            {
                // Cache the answer to avoid pounding on the registry
                if (isOurPropertyHandlerRegistered == null)
                    using (RegistryKey key = Registry.ClassesRoot.OpenSubKey(@"CLSID\" + OurPropertyHandlerGuid, false))
                    {
                        isOurPropertyHandlerRegistered = (key != null);
                    }

                return (bool) isOurPropertyHandlerRegistered;
            }
        }

        public static bool IsOurContextHandlerRegistered
        {
            get
            {
                // Cache the answer to avoid pounding on the registry
                if (isOurContextHandlerRegistered == null)
                    using (RegistryKey key = Registry.ClassesRoot.OpenSubKey(@"CLSID\" + OurContextHandlerGuid, false))
                    {
                        isOurContextHandlerRegistered = (key != null);
                    }

                return (bool) isOurContextHandlerRegistered;
            }
        }

        public void RecordPropertyHandler(string handlerGuid, string handlerChainedGuid)
        {
            if (handlerGuid == OurPropertyHandlerGuid)
            {
                if (handlerChainedGuid != null)
                    RecordPropertyHandler(HandlerState.Chained, handlerChainedGuid, GetHandlerTitle(handlerChainedGuid));
                else
                    RecordPropertyHandler(HandlerState.Ours, null, null);
            }
            else
                RecordPropertyHandler(HandlerState.Foreign, handlerGuid, GetHandlerTitle(handlerGuid));
        }

        public void IdentifyCurrentProfile()
        {
            Profile = GetCurrentProfileIfKnown();
        }

        public Profile GetDefaultCustomProfile()
        {
            var p = new Profile();

            // Prefer configuration held in SystemFileAssociations
            using (RegistryKey target = GetSystemFileAssociationsProfileKey(false))
            {
                if (!GetExistingProgidRegistryEntries(target, p))
                {
                    // But if there is none, try the extension entry in HKCR
                    using (RegistryKey target2 = GetHKCRProfileKey(false))
                    {
                        GetExistingProgidRegistryEntries(target2, p);
                    }
                }
            }
            return p;
        }
       
        // Setup the registry entries required when an extension is given our handler
        public bool SetupHandlerForExtension(Profile selectedProfile)
        {
            if (PropertyHandlerState != HandlerState.None && PropertyHandlerState != HandlerState.Foreign)
                return false;

            if (!IsElevated)
                return false;

            Profile profile;

            // Work out what profile to set up
            if (PropertyHandlerState == HandlerState.None)
                profile = selectedProfile;
            else
            {
                // We use a custom profile with the same name as the extension
                profile = State.CustomProfiles.FirstOrDefault(p => p.Name == Name);

                // If it doesn't already exist, create it
                if (profile == null)
                {
                    profile = new Profile { Name = this.Name, State = this.State };
                    State.CustomProfiles.Add(profile);
                }
                // If we're recycling the profile, clone it to preserve its original values
                else if (profile == selectedProfile)
                    selectedProfile = selectedProfile.CreateClone();
            }

            // Find the key for the extension in HKEY_CLASSES_ROOT
            using (RegistryKey target = GetHKCRProfileKey(true))
            {
                // We used to place entries on this key, but no longer do, because such keys can be shared,
                // and we only want to affect a specific extension
                // We still have to hide any existing entries, because otherwise they would take priority, but fortunately they do not often occur
                // If there are entries, we read them into the profile that we are setting up
                if (PropertyHandlerState == HandlerState.Foreign)
                    GetAndHidePreExistingProgidRegistryEntries(target, profile);
            }

            // Now we only update the extension specific area
            using (RegistryKey target = GetSystemFileAssociationsProfileKey(true))
            {
                if (PropertyHandlerState == HandlerState.Foreign)
                {
                    GetAndHidePreExistingProgidRegistryEntries(target, profile);

                    // Merge the selected profile into any entries that came with the foreign handler
                    profile.MergeFrom(selectedProfile);
                }

                if (!SetupProgidRegistryEntries(target, profile))
                    return false;
            }

            if (PropertyHandlerState == HandlerState.Foreign)
            {
                // Write the updated custom profile information back to the store
                State.StoreSavedState();
            }              

            // Now, add the property handler extension key
            // Watch out for 32/64 bit issues here, as the 32-bit and 64-bit values of these are separate and isolated on 64-bit Windows,
            // the 32-bit value being under SOFTWARE\Wow6432Node.  Thus a 64-bit manager is needed to set up a 64-bit handler
            using (RegistryKey handlers = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\PropertySystem\PropertyHandlers", true))
            {
                using (RegistryKey handler = handlers.CreateSubKey(Name))
                {
                    if (PropertyHandlerState == HandlerState.None)
                    {
                        handler.SetValue(null, OurPropertyHandlerGuid);
                        this.RecordPropertyHandler(HandlerState.Ours, null, null);
                    }
                    else // Foreign
                    {
                        handler.SetValue(null, OurPropertyHandlerGuid);
                        handler.SetValue(ChainedValueName, PropertyHandlerGuid);
                        this.RecordPropertyHandler(HandlerState.Chained, PropertyHandlerGuid, PropertyHandlerTitle);
                    }
                }
            }
            this.Profile = profile;

            State.HasChanged = true;

            return true;
        }

        // Update the registry settings for an extension when the Full details or Preview details in a profile are changed
        public bool UpdateProfileSettingsForExtension(Profile profile)
        {
            if (PropertyHandlerState != HandlerState.Ours && PropertyHandlerState != HandlerState.Chained)
                return false;

            if (!IsElevated)
                return false;

            using (RegistryKey target = GetSystemFileAssociationsProfileKey(true))
            {
                return SetupProfileDetailEntries(target, profile);
            }
        }

        private bool GetExistingProgidRegistryEntries(RegistryKey target, Profile profile)
        {
            if (target == null)
                return false;

            var val = target.GetValue(FullDetailsValueName);
            if (val != null)
                profile.FullDetailsString = val as string;

            val = target.GetValue(InfoTipValueName);
            if (val != null)
                profile.InfoTipString = val as string;

            val = target.GetValue(PreviewDetailsValueName);
            if (val != null)
                profile.PreviewDetailsString = val as string;


            return val != null;
        }

        private void GetAndHidePreExistingProgidRegistryEntries(RegistryKey target, Profile profile)
        {
            if (target == null)
                return;

            // We only have to hide existing Profile Detail entries, as context menus are additive
            var val = target.GetValue(FullDetailsValueName);
            if (val != null)
            {
                target.SetValue(OldFullDetailsValueName, val);
                target.DeleteValue(FullDetailsValueName);
                profile.FullDetailsString = val as string;
            }

            val = target.GetValue(InfoTipValueName);
            if (val != null)
            {
                target.SetValue(OldInfoTipValueName, val);
                target.DeleteValue(InfoTipValueName);
                profile.InfoTipString = val as string;
            }

            val = target.GetValue(PreviewDetailsValueName);
            if (val != null)
            {
                target.SetValue(OldPreviewDetailsValueName, val);
                target.DeleteValue(PreviewDetailsValueName);
                profile.PreviewDetailsString = val as string;
            }
        }

        private bool SetupProgidRegistryEntries(RegistryKey target, Profile profile)
        {
            if (target == null)
                return false;

            SetupProfileDetailEntries(target, profile);

            // Set up the eontext handler, if registered
            if (IsOurContextHandlerRegistered)
            {
                using (RegistryKey keyShellEx = target.CreateSubKey(ShellExKeyName))
                {
                    using (RegistryKey keyCMH = keyShellEx.CreateSubKey(ContextMenuHandlersKeyName))
                    {
                        using (RegistryKey keyHandler = keyCMH.CreateSubKey(ContextHandlerKeyName))
                        {
                            keyHandler.SetValue(null, OurContextHandlerGuid);
                        }
                    }
                }
            }

            return true;
        }

        private bool SetupProfileDetailEntries(RegistryKey target, Profile profile)
        {
            if (target == null)
                return false;

            // Update the info tip and full and preview details keys
            target.SetValue(FullDetailsValueName, profile.FullDetailsString);
            target.SetValue(InfoTipValueName, profile.InfoTipString);
            target.SetValue(PreviewDetailsValueName, profile.PreviewDetailsString);

            // If this is a custom profile, update our private key that records its name
            if (!profile.IsReadOnly)
                target.SetValue(FileMetaCustomProfileValueName, profile.Name);

            return true;
        }

        private Profile GetCurrentProfileIfKnown()
        {
            if (PropertyHandlerState != HandlerState.Ours && PropertyHandlerState != HandlerState.Chained)
                return null;

            // Find the key for the extension 
            string pd;
            string cp;
            using (RegistryKey target = GetSystemFileAssociationsProfileKey(false))
            {
                if (target == null)
                    return null;

                // Try to identify the profile
                pd = (string)target.GetValue(PreviewDetailsValueName);
                cp = (string)target.GetValue(FileMetaCustomProfileValueName);
            }

            // If our private key recording the custom profile used is present, go with that
            if (cp != null)
                return State.CustomProfiles.Where(p => p.Name == cp).FirstOrDefault();

            // Otherwise, tried to match the preview details against one of the built-in values
            else if (pd != null)
                return State.BuiltInProfiles.Where(p => p.PreviewDetailsString == pd).FirstOrDefault();

            return null;
        }

        public void RemoveHandlerFromExtension()
        {
            if (PropertyHandlerState != HandlerState.Ours && PropertyHandlerState != HandlerState.Chained)
                return;

            // Now find the key for the extension in HKEY_CLASSES_ROOT
            using (RegistryKey target = GetHKCRProfileKey(true))
            {
                // Tolerate the case where the extension has been removed since we set up a handler for it
                // We still do this even though no longer write entries here so as to be sure to clean up after earlier versions,
                // and to restore pre-existing entries if we had to hide them
                if (target != null)
                    RemoveProgidRegistryEntries(target);
            }

            // Now go after the settings in SystemFileAssociations
            // We may have created the key for the extension, but we can't tell, so we just remove the values
            using (RegistryKey target = GetSystemFileAssociationsProfileKey(true))
            {
                if (target != null)
                    RemoveProgidRegistryEntries(target);
            }

            // Now, remove the handler extension key
            // Watch out for 32/64 bit issues here, as the 32-bit and 64-bit values of these are separate and isolated on 64-bit Windows,
            // the 32-bit value being under SOFTWARE\Wow6432Node.  Thus a 64-bit manager is needed to set up a 64-bit handler
            using (RegistryKey handlers = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\PropertySystem\PropertyHandlers", true))
            {
                if (PropertyHandlerState == HandlerState.Ours)
                {
                    handlers.DeleteSubKey(Name);
                    this.RecordPropertyHandler(HandlerState.None, null, null);
                    this.Profile = null;
                }
                else  // Chained
                {
                    using (RegistryKey handler = handlers.OpenSubKey(Name, true))
                    {
                        handler.SetValue(null, PropertyHandlerGuid);
                        handler.DeleteValue(ChainedValueName);
                        this.RecordPropertyHandler(HandlerState.Foreign, PropertyHandlerGuid, PropertyHandlerTitle);
                        this.Profile = null;
                    }
                }
            }

            State.HasChanged = true;
        }

        private void RemoveProgidRegistryEntries(RegistryKey target)
        {
            // tolerate missing entries
            if (target == null)
                return;

            target.DeleteValue(FullDetailsValueName, false);
            target.DeleteValue(InfoTipValueName, false);
            target.DeleteValue(PreviewDetailsValueName, false);
            target.DeleteValue(FileMetaCustomProfileValueName, false);

            UnhidePreExistingProgidRegistryEntries(target);

            // Always have a go at removing the context handler setup, even if the handler is not registered
            // There might be entries lying around from when it was
            using (RegistryKey keyShellEx = target.OpenSubKey(ShellExKeyName, true))
            {
                if (keyShellEx != null)
                {
                    using (RegistryKey keyCMH = keyShellEx.OpenSubKey(ContextMenuHandlersKeyName, true))
                    {
                        if (keyCMH != null)
                        {
                            keyCMH.DeleteSubKey(ContextHandlerKeyName, false);
                            if (keyCMH.SubKeyCount == 0)
                            {
                                keyCMH.Close();
                                keyShellEx.DeleteSubKey(ContextMenuHandlersKeyName);
                            }
                            else
                                keyCMH.Close();
                        }
                    }
                    // Only remove ShellEx if it is now completely empty
                    if (keyShellEx.SubKeyCount == 0 && keyShellEx.ValueCount == 0)
                    {
                        keyShellEx.Close();
                        target.DeleteSubKey(ShellExKeyName);
                    }
                    else
                        keyShellEx.Close();
                }
            }

            target.Close();
        }

        private void UnhidePreExistingProgidRegistryEntries(RegistryKey target)
        {
            if (target == null)
                return;

            // We only have to unhide existing Profile Detail entries, as that is all we hide
            var val = target.GetValue(OldFullDetailsValueName);
            if (val != null)
            {
                target.SetValue(FullDetailsValueName, val);
                target.DeleteValue(OldFullDetailsValueName);
            }

            val = target.GetValue(OldInfoTipValueName);
            if (val != null)
            {
                target.SetValue(InfoTipValueName, val);
                target.DeleteValue(OldInfoTipValueName);
            }

            val = target.GetValue(OldPreviewDetailsValueName);
            if (val != null)
            {
                target.SetValue(PreviewDetailsValueName, val);
                target.DeleteValue(OldPreviewDetailsValueName);
            }
        }

        private RegistryKey GetHKCRProfileKey(bool bWritable)
        {
            // Find the key for the extension in HKEY_CLASSES_ROOT
            string progid;
            using (RegistryKey ext = Registry.ClassesRoot.OpenSubKey(Name, false))
            {
                if (ext == null)
                    return null;

                progid = (string)ext.GetValue(null);
            }

            RegistryKey target;
            string targetName;

            if (progid != null && progid.Length > 0)
                targetName = progid;
            else
                targetName = Name;

            try
            {
                target = Registry.ClassesRoot.OpenSubKey(progid, bWritable);
            }
            catch (System.Security.SecurityException e)
            {
                MessageBox.Show(String.Format(LocalizedMessages.NoRegistryPermission, targetName), LocalizedMessages.ProfileError);
                target = null;
            }

            return target;
        }

        private RegistryKey GetSystemFileAssociationsProfileKey(bool bWritable)
        {
            using (RegistryKey assoc = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Classes\SystemFileAssociations", bWritable))
            {
                if (assoc == null)
                    return null;

                if (bWritable)
                    return assoc.CreateSubKey(Name);
                else
                    return assoc.OpenSubKey(Name, false);
            }
        }

        private void RecordPropertyHandler(HandlerState propertyHandlerState, string guid, string title)
        {
            PropertyHandlerState = propertyHandlerState;
            PropertyHandlerGuid = guid;
            PropertyHandlerTitle = title;

            OnPropertyChanged("PropertyHandlerDisplay");
            OnPropertyChanged("Weight");
        }

        private string GetHandlerTitle(string handlerGuid)
        {
            string handlerTitle = null;
            var cls = Registry.ClassesRoot.OpenSubKey(@"CLSID\" + handlerGuid);
            if (cls != null)
            {
                handlerTitle = (string)cls.GetValue(null);
                if (handlerTitle == null)
                {
                    // No name - check for shell handlers
                    if (handlerGuid == "{66742402-F9B9-11D1-A202-0000F81FEDEE}" ||
                        handlerGuid == "{0AFCCBA6-BF90-4A4E-8482-0AC960981F5B}")
                        handlerTitle = "Windows Shell";
                    else
                    {
                        // Else resort to the dll path
                        var ps = cls.OpenSubKey("InProcServer32");
                        if (ps != null)
                            handlerTitle = (string)ps.GetValue(null);
                    }
                }
            }

            return handlerTitle;
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
