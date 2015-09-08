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
    public class Extension : INotifyPropertyChanged
    {
        const string OurPropertyHandlerTitle = "File Meta Property Handler";
        const string OurPropertyHandlerGuid64 = "{D06391EE-2FEB-419B-9667-AD160D0849F3}";
        const string OurPropertyHandlerGuid32 = "{60211757-EF87-465e-B6C1-B37CF98295F9}";
        const string OurContextHandlerGuid64 = "{28D14D00-2D80-4956-9657-9D50C8BB47A5}";
        const string OurContextHandlerGuid32 = "{DA38301B-BE91-4397-B2C8-E27A0BD80CC5}";

        const string FullDetailsValueName = "FullDetails";
        const string PreviewDetailsValueName = "PreviewDetails";
        const string FileMetaCustomProfileValueName = "FileMetaCustomProfile";
        const string ShellExKeyName = "ShellEx";
        const string ContextMenuHandlersKeyName = "ContextMenuHandlers";
        const string ContextHandlerKeyName = "FileMetadata";

        private string propertyHandlerGuid = null;  // hold as string because Guid.Parse() needs .Net 4
        private string propertyHandlerTitle;
        private static Nullable<bool> isOurPropertyHandlerRegistered = null;
        private static Nullable<bool> isOurContextHandlerRegistered = null;

        private Profile profile = null;  // The profile currently applied to this extension, if any

        public string Name { get; set; }
        public State State { get; set; }
        public Profile Profile { get { return profile; } set { profile = value; } }
        public FontWeight Weight { get { return OurHandler ? FontWeights.ExtraBold : (ForeignHandler ? FontWeights.ExtraLight : FontWeights.Normal); } }
        public string PropertyHandlerGuid { get { return propertyHandlerGuid; } }
        public string PropertyHandlerTitle { get { return propertyHandlerTitle; } }
        public string PropertyHandlerNow 
        {
            get 
            { 
                return String.Format(LocalizedMessages.PropertyHandlerCurrent, Name,
                    (propertyHandlerTitle != null ? propertyHandlerTitle : (propertyHandlerGuid != null ? propertyHandlerGuid : LocalizedMessages.PropertyHandlerNone))); 
            } 
        }

        public bool HasHandler { get { return propertyHandlerGuid != null; } }
        public bool ForeignHandler { get { return propertyHandlerGuid != null && propertyHandlerGuid != OurPropertyHandlerGuid; } }
        public bool OurHandler { get { return propertyHandlerGuid != null && propertyHandlerGuid == OurPropertyHandlerGuid; } }

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
                    isOurPropertyHandlerRegistered = (Registry.ClassesRoot.OpenSubKey(@"CLSID\" + OurPropertyHandlerGuid, false) != null);

                return (bool) isOurPropertyHandlerRegistered;
            }
        }

        public static bool IsOurContextHandlerRegistered
        {
            get
            {
                // Cache the answer to avoid pounding on the registry
                if (isOurContextHandlerRegistered == null)
                    isOurContextHandlerRegistered = (Registry.ClassesRoot.OpenSubKey(@"CLSID\" + OurContextHandlerGuid, false) != null);

                return (bool) isOurContextHandlerRegistered;
            }
        }

        public void RecordPropertyHandler (string guid, string title)
        {
            propertyHandlerGuid = guid;
            propertyHandlerTitle = title;

            OnPropertyChanged("PropertyHandlerTitle");
            OnPropertyChanged("PropertyHandlerNow");
            OnPropertyChanged("Weight");
        }

        public void IdentifyCurrentProfile()
        {
            Profile = GetCurrentProfileIfKnown();
        }
       
        // Setup the registry entries required when an extension is given our handler
        public bool SetupHandlerForExtension(Profile profile)
        {
            if (ForeignHandler || OurHandler)
                return false;

            if (!IsElevated)
                return false;

            // Find the key for the extension in HKEY_CLASSES_ROOT
            var ext = Registry.ClassesRoot.OpenSubKey(Name,false);
            if (ext == null)
                return false;

            string progid = (string)ext.GetValue(null);

            string targetName;
            RegistryKey target;
            if (progid != null && progid.Length > 0)
                targetName = progid;
            else
                targetName = Name;

            try
            {
                target = Registry.ClassesRoot.OpenSubKey(targetName, true);
            }
            catch (System.Security.SecurityException e)
            {
                MessageBox.Show(String.Format(LocalizedMessages.NoRegistryPermission, targetName), LocalizedMessages.ProfileError);
                return false;
            }

            if (!SetupProgidRegistryEntries(target, profile))
                return false;

             // Sometimes, the details don't seem to get picked up from ProgID, so put the same stuff in what is supposed to be a lower priority location
            // to catch some more cases
            var assoc = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Classes\SystemFileAssociations", true);
            target = assoc.CreateSubKey(Name);

            if (!SetupProgidRegistryEntries(target, profile))
                return false;

            // Now, add the property handler extension key
            // Watch out for 32/64 bit issues here, as the 32-bit and 64-bit values of these are separate and isolated on 64-bit Windows,
            // the 32-bit value being under SOFTWARE\Wow6432Node.  Thus a 64-bit manager is needed to set up a 64-bit handler
            var handlers = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\PropertySystem\PropertyHandlers", true);
            var handler = handlers.CreateSubKey(Name);
            handler.SetValue(null, OurPropertyHandlerGuid);
            this.RecordPropertyHandler(OurPropertyHandlerGuid, OurPropertyHandlerTitle);
            this.Profile = profile;
            State.HasChanged = true;

            return true;
        }

        // Update the registry settings for an extension when the Full details or Preview details in a profile are changed
        public bool UpdateProfileSettingsForExtension(Profile profile)
        {
            if (!OurHandler)
                return false;

            if (!IsElevated)
                return false;

            // Find the key for the extension in HKEY_CLASSES_ROOT
            var ext = Registry.ClassesRoot.OpenSubKey(Name, false);
            if (ext == null)
                return false;

            string progid = (string)ext.GetValue(null);

            string targetName;
            RegistryKey target;
            if (progid != null && progid.Length > 0)
                targetName = progid;
            else
                targetName = Name;

            try
            {
                target = Registry.ClassesRoot.OpenSubKey(targetName, true);
            }
            catch (System.Security.SecurityException e)
            {
                MessageBox.Show(String.Format(LocalizedMessages.NoRegistryPermission, targetName), LocalizedMessages.ProfileError);
                return false;
            }

            if (!SetupProfileDetailEntries(target, profile))
                return false;

            // Sometimes, the details don't seem to get picked up from ProgID, so put the same stuff in what is supposed to be a lower priority location
            // to catch some more cases
            var assoc = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Classes\SystemFileAssociations", true);
            target = assoc.CreateSubKey(Name);

            if (!SetupProfileDetailEntries(target, profile))
                return false;

            return true;
        }

        private bool SetupProgidRegistryEntries (RegistryKey target, Profile profile)
        {
            if (target == null)
                return false;

            SetupProfileDetailEntries(target, profile);

            // Set up the eontext handler, if registered
            if (IsOurContextHandlerRegistered)
            {
                var keyShellEx = target.CreateSubKey(ShellExKeyName);
                var keyCMH = keyShellEx.CreateSubKey(ContextMenuHandlersKeyName);
                var keyHandler = keyCMH.CreateSubKey(ContextHandlerKeyName);
                keyHandler.SetValue(null, OurContextHandlerGuid);
            }

            return true;
        }

        private bool SetupProfileDetailEntries(RegistryKey target, Profile profile)
        {
            if (target == null)
                return false;

            // Update the full and preview details keys
            target.SetValue(FullDetailsValueName, profile.FullDetailsString);
            target.SetValue(PreviewDetailsValueName, profile.PreviewDetailsString);

            // If this is a custom profile, update our private key that records its name
            if (!profile.IsReadOnly)
                target.SetValue(FileMetaCustomProfileValueName, profile.Name);

            return true;
        }

        private Profile GetCurrentProfileIfKnown()
        {
            if (!OurHandler)
                return null;

            // Find the key for the extension in HKEY_CLASSES_ROOT
            var ext = Registry.ClassesRoot.OpenSubKey(Name, false);
            if (ext == null)
                return null;

            string progid = (string)ext.GetValue(null);

            RegistryKey target;
            if (progid != null && progid.Length > 0)
                target = Registry.ClassesRoot.OpenSubKey(progid, false);
            else
                target = Registry.ClassesRoot.OpenSubKey(Name, false);

            if (target == null)
                return null;

            // Try to identify the profile
            string pd = (string) target.GetValue(PreviewDetailsValueName);
            string cp = (string)target.GetValue(FileMetaCustomProfileValueName);

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
            if (!OurHandler)
                return;

            // Now find the key for the extension in HKEY_CLASSES_ROOT
            RegistryKey target;
            var ext = Registry.ClassesRoot.OpenSubKey(Name, false);

            // Tolerate the case where the extension has been removed since we set up a handler for it
            if (ext != null)
            {
                string progid = (string)ext.GetValue(null);
                string targetName;

                if (progid != null && progid.Length > 0)
                    targetName = progid;
                else
                    targetName = Name;

                try
                {
                    target = Registry.ClassesRoot.OpenSubKey(targetName, true);
                }
                catch (System.Security.SecurityException e)
                {
                    MessageBox.Show(String.Format(LocalizedMessages.NoRegistryPermission, targetName), LocalizedMessages.ProfileError);
                    return;
                }

                RemoveProgidRegistryEntries(target);
            }

            // Now go after the settings in SystemFileAssociations
            // We may have created the key for the extension, but we can't tell, so we just remove the values
            var assoc = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Classes\SystemFileAssociations", true);
            target = assoc.OpenSubKey(Name, true);

            RemoveProgidRegistryEntries(target);

            // Now, remove the handler extension key
            // Watch out for 32/64 bit issues here, as the 32-bit and 64-bit values of these are separate and isolated on 64-bit Windows,
            // the 32-bit value being under SOFTWARE\Wow6432Node.  Thus a 64-bit manager is needed to set up a 64-bit handler
            var handlers = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\PropertySystem\PropertyHandlers", true);
            if (handlers != null)
                handlers.DeleteSubKey(Name);
            this.RecordPropertyHandler(null, null);
            this.Profile = null;
            State.HasChanged = true;
        }

        private void RemoveProgidRegistryEntries(RegistryKey target)
        {
            // tolerate missing entries
            if (target == null)
                return;

            target.DeleteValue(FullDetailsValueName, false);
            target.DeleteValue(PreviewDetailsValueName, false);
            target.DeleteValue(FileMetaCustomProfileValueName, false);

            // Always have a go at removing the context handler setup, even if the handler is not registered
            // There might be entries lying around from when it was
            var keyShellEx = target.OpenSubKey(ShellExKeyName, true);
            if (keyShellEx != null)
            {
                var keyCMH = keyShellEx.OpenSubKey(ContextMenuHandlersKeyName, true);
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
                // Only remove ShellEx if it is now completely empty
                if (keyShellEx.SubKeyCount == 0 && keyShellEx.ValueCount == 0)
                {
                    keyShellEx.Close();
                    target.DeleteSubKey(ShellExKeyName);
                }
                else
                    keyShellEx.Close();
            }

            target.Close();
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
