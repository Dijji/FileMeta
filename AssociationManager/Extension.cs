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
using AssociationMessages;

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

#if CmdLine
#else
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
#endif

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
            else if (handlerGuid != null)
                RecordPropertyHandler(HandlerState.Foreign, handlerGuid, GetHandlerTitle(handlerGuid));
            else
                RecordPropertyHandler(HandlerState.None, null, null);
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
        public void SetupHandlerForExtension(Profile selectedProfile, bool createMergedProfile)
        {
            if (!IsElevated)
                throw new AssocMgrException
                {
                    Description = LocalizedMessages.AdminPrivilegesNeeded,
                    Exception = null,
                    ErrorCode = WindowsErrorCodes.ERROR_ACCESS_DENIED
                };

            if (PropertyHandlerState != HandlerState.None && PropertyHandlerState != HandlerState.Foreign)
                throw new AssocMgrException
                {
                    Description = String.Format(LocalizedMessages.ExtensionAlreadyHasHandler, Name),
                    Exception = null,
                    ErrorCode = WindowsErrorCodes.ERROR_INVALID_PARAMETER
                };

            Profile profile;

            // Work out what profile to set up
            if (PropertyHandlerState == HandlerState.None || !createMergedProfile)
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
                // If there are entries, and we are merging, we read them into the new profile, in case this is the only place they occur
                if (PropertyHandlerState == HandlerState.Foreign)
                    GetAndHidePreExistingProgidRegistryEntries(target, createMergedProfile ? profile : null);
            }

            // Now we only update the extension specific area
            using (RegistryKey target = GetSystemFileAssociationsProfileKey(true))
            {
                if (PropertyHandlerState == HandlerState.Foreign)
                {
                    if (createMergedProfile)
                    {
                        GetAndHidePreExistingProgidRegistryEntries(target, profile);

                        // Merge the selected profile into any entries that came with the foreign handler
                        profile.MergeFrom(selectedProfile);
                    }
                    else
                        GetAndHidePreExistingProgidRegistryEntries(target, null); 
                }

                SetupProgidRegistryEntries(target, profile);
            }

            if (PropertyHandlerState == HandlerState.Foreign)
            {
                // Write the updated custom profile information back to the store
                State.StoreUpdatedProfile(profile);
            }              

#if x64
            // On 64-bit machines, set up the 32-bit property handler, so that 32-bit applications can also access our properties
            using (RegistryKey handlers = RegistryExtensions.OpenBaseKey(RegistryHive.LocalMachine, RegistryExtensions.RegistryHiveType.X86).
                                            OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\PropertySystem\PropertyHandlers", true))
            {
                using (RegistryKey handler = handlers.CreateSubKey(Name))
                {
                    if (PropertyHandlerState == HandlerState.None)
                    {
                        var temp = handler.GetValue(null); 
                        // In the case where there is no 64-bit handler, but there is a 32-bit handler, leave it alone
                        if (temp == null)
                            handler.SetValue(null, OurPropertyHandlerGuid32);
                    }
                    else // Foreign
                    {
                        var temp = handler.GetValue(null);
                        handler.SetValue(null, OurPropertyHandlerGuid32);
                        handler.SetValue(ChainedValueName, temp);
                    }
                }
            }
#endif 
            // Now, add the main handler extension key, which is 32- or 64-bit, depending on how we were built
            // The 32-bit and 64-bit values of these are separate and isolated on 64-bit Windows,
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
        }

        // Update the registry settings for an extension when the Full details or Preview details in a profile are changed
        public void UpdateProfileSettingsForExtension(Profile profile)
        {
            if (!IsElevated)
                throw new AssocMgrException
                {
                    Description = LocalizedMessages.AdminPrivilegesNeeded,
                    Exception = null,
                    ErrorCode = WindowsErrorCodes.ERROR_ACCESS_DENIED
                };

            if (PropertyHandlerState != HandlerState.Ours && PropertyHandlerState != HandlerState.Chained)
                return;

            using (RegistryKey target = GetSystemFileAssociationsProfileKey(true))
            {
                SetupProfileDetailEntries(target, profile);
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

        // If the caller does not need the existing values, it should pass null for the profile
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
                if (profile != null)
                    profile.FullDetailsString = val as string;
            }

            val = target.GetValue(InfoTipValueName);
            if (val != null)
            {
                target.SetValue(OldInfoTipValueName, val);
                target.DeleteValue(InfoTipValueName);
                if (profile != null)
                    profile.InfoTipString = val as string;
            }

            val = target.GetValue(PreviewDetailsValueName);
            if (val != null)
            {
                target.SetValue(OldPreviewDetailsValueName, val);
                target.DeleteValue(PreviewDetailsValueName);
                if (profile != null)
                    profile.PreviewDetailsString = val as string;
            }
        }

        private void SetupProgidRegistryEntries(RegistryKey target, Profile profile)
        {
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
        }

        private void SetupProfileDetailEntries(RegistryKey target, Profile profile)
        {
            // Update the info tip and full and preview details keys
            target.SetValue(FullDetailsValueName, profile.FullDetailsString);
            target.SetValue(InfoTipValueName, profile.InfoTipString);
            target.SetValue(PreviewDetailsValueName, profile.PreviewDetailsString);

            // If this is a custom profile, update our private key that records its name
            if (!profile.IsReadOnly)
                target.SetValue(FileMetaCustomProfileValueName, profile.Name);
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

        // Check to see if the registry entries are down level i.e. 1.3 or 1.4, and so need refreshing
        public bool IsRefreshRequired()
        {
            // Find the key for the extension in HKEY_CLASSES_ROOT
            using (RegistryKey target = GetHKCRProfileKey(true))
            {
                // Refresh is required if old school FullDetails value is present 
                if (target != null && target.GetValue(FullDetailsValueName) != null)
                    return true;
            }
#if x64
            // On 64-bit machines, check that we have a 32-bit property handler, as we now set up both 
            using (RegistryKey handlers = RegistryExtensions.OpenBaseKey(RegistryHive.LocalMachine, RegistryExtensions.RegistryHiveType.X86).
                                            OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\PropertySystem\PropertyHandlers", true))
            {
                using (RegistryKey handler = handlers.OpenSubKey(Name, false))
                {
                    // It is important to get the criteria right here, as we don't want to
                    // return a positive unless removing the handler and setting it up again
                    // will resolve the situation

                    // If there is no handler, refresh is required
                    if (handler == null || handler.GetValue(null) == null)
                        return true;
                    // Also refresh if we are chained, and no original property handler is recorded (even if null)
                    else if (PropertyHandlerState == HandlerState.Chained && !handler.GetValueNames().Contains(ChainedValueName))
                        return true;
                }
            }
#endif 
            return false;
        }

        public void RemoveHandlerFromExtension()
        {
            if (!IsElevated)
                throw new AssocMgrException
                {
                    Description = LocalizedMessages.AdminPrivilegesNeeded,
                    Exception = null,
                    ErrorCode = WindowsErrorCodes.ERROR_ACCESS_DENIED
                };

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

#if x64
            // On 64-bit machines, remove the 32-bit property handler, as we set up both 
            using (RegistryKey handlers = RegistryExtensions.OpenBaseKey(RegistryHive.LocalMachine, RegistryExtensions.RegistryHiveType.X86).
                                            OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\PropertySystem\PropertyHandlers", true))
            {
                if (PropertyHandlerState == HandlerState.Ours)
                {
                    bool delete = false;
                    using (RegistryKey handler = handlers.OpenSubKey(Name, true))
                    {
                        if (handler != null)
                        {
                            // Only delete the sub key if it points to our handler
                            var temp = handler.GetValue(null) as string;
                            delete = (temp != null && temp == OurPropertyHandlerGuid32);
                        }
                    }
                    // Delete needs to happen after we have released the registry key
                    if (delete)
                        handlers.DeleteSubKey(Name);
                }
                else  // Chained
                {
                    using (RegistryKey handler = handlers.OpenSubKey(Name, true))
                    {
                        if (handler != null)
                        {
                            // Allow for the case where the chained value exists but is empty
                            if (handler.GetValueNames().Contains(ChainedValueName))
                            {
                                var temp = handler.GetValue(ChainedValueName);
                                handler.SetValue(null, temp);
                                handler.DeleteValue(ChainedValueName);
                            }
                        }
                    }
                }
            }
#endif 
            // Now, remove the main handler extension key, which is 32- or 64-bit, depending on how we were built
            // The 32-bit and 64-bit values of these are separate and isolated on 64-bit Windows,
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

            using (RegistryKey extKey = Registry.ClassesRoot.OpenSubKey(Name, false))
            {
                if (extKey == null)
                {
                    if (bWritable)
                        progid = null;
                    else
                        return null;
                }
                else
                    progid = (string)extKey.GetValue(null);
            }

            string targetName;

            if (progid != null && progid.Length > 0)
                targetName = progid;
            else
                targetName = Name;

            try
            {
                if (bWritable)
                    return Registry.ClassesRoot.CreateSubKey(targetName);
                else
                    return Registry.ClassesRoot.OpenSubKey(targetName, false);
            }
            catch (System.Security.SecurityException e)
            {
                throw new AssocMgrException
                {
                    Description = String.Format(LocalizedMessages.NoRegistryPermission, targetName),
                    Exception = e,
                    ErrorCode = WindowsErrorCodes.ERROR_ACCESS_DENIED
                };
            }
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
