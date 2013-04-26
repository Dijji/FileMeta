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
        const string ShellExKeyName = "ShellEx";
        const string ContextMenuHandlersKeyName = "ContextMenuHandlers";
        const string ContextHandlerKeyName = "FileMetadata";

        private string propertyHandlerGuid = null;  // hold as string because Guid.Parse() needs .Net 4
        private string propertyHandlerTitle;
        private static Nullable<bool> isOurPropertyHandlerRegistered = null;
        private static Nullable<bool> isOurContextHandlerRegistered = null;


        public string Name { get; set; }
        public State State { get; set; }
        public FontStyle Style { get { return ForeignHandler ? FontStyles.Italic : FontStyles.Normal; } }
        public FontWeight Weight { get { return OurHandler ? FontWeights.ExtraBold : (ForeignHandler ? FontWeights.ExtraLight : FontWeights.Normal); } }
        public string PropertyHandlerGuid { get { return propertyHandlerGuid; } }
        public string PropertyHandlerTitle { get { return propertyHandlerTitle; } }
        public bool CanAddPropertyHandlerEtc { get { return IsOurPropertyHandlerRegistered && propertyHandlerGuid == null && State.SelectedProfile != null; } }
        public bool CanRemovePropertyHandlerEtc { get { return OurHandler; } }
        public string PropertyHandlerNow 
        {
            get 
            { 
                return String.Format(LocalizedMessages.PropertyHandlerCurrent, Name,
                    (propertyHandlerTitle != null ? propertyHandlerTitle : (propertyHandlerGuid != null ? propertyHandlerGuid : LocalizedMessages.PropertyHandlerNone))); 
            } 
        }

        private bool ForeignHandler { get { return propertyHandlerGuid != null && propertyHandlerGuid != OurPropertyHandlerGuid; } }
        private bool OurHandler { get { return propertyHandlerGuid != null && propertyHandlerGuid == OurPropertyHandlerGuid; } }

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
            OnPropertyChanged("Style");
            OnPropertyChanged("Weight");
            OnPropertyChanged("CanAddPropertyHandlerEtc");
            OnPropertyChanged("CanRemovePropertyHandlerEtc");
        }
       
        public bool SetupHandlerForExtension(Profile profile)
        {
            if (ForeignHandler || OurHandler)
                return false;

            if (!IsElevated)
                return false;

            // First, add the property handler extension key
            // Watch out for 32/64 bit issues here, as the 32-bit and 64-bit values of these are separate and isolated on 64-bit Windows,
            // the 32-bit value being under SOFTWARE\Wow6432Node.  Thus a 64-bit manager is needed to set up a 64-bit handler
            var handlers = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\PropertySystem\PropertyHandlers", true);
            var handler = handlers.CreateSubKey(Name);
            handler.SetValue(null, OurPropertyHandlerGuid);
            this.RecordPropertyHandler (OurPropertyHandlerGuid, OurPropertyHandlerTitle);

            // Now find the key for the extension in HKEY_CLASSES_ROOT
            var ext = Registry.ClassesRoot.OpenSubKey(Name,false);
            if (ext == null)
                return false;

            string progid = (string)ext.GetValue(null);

            RegistryKey target;
            if (progid.Length > 0)
                target = Registry.ClassesRoot.OpenSubKey(progid, true);
            else
                target = Registry.ClassesRoot.OpenSubKey(Name, true);

            if (!SetupProgidRegistryEntries(target, profile))
                return false;

             // Sometimes, the details don't seem to get picked up from ProgID, so put the same stuff in what is supposed to be a lower priority location
            // to catch some more cases
            var assoc = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Classes\SystemFileAssociations", true);
            target = assoc.CreateSubKey(Name);

            return SetupProgidRegistryEntries(target, profile);
        }

        private bool SetupProgidRegistryEntries (RegistryKey target, Profile profile)
        {
            if (target == null)
                return false;

            target.SetValue(FullDetailsValueName, profile.FullDetailsString);
            target.SetValue(PreviewDetailsValueName, profile.PreviewDetailsString);

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

        public void SelectCurrentProfileIfKnown()
        {
            if (!OurHandler)
                return;

            // Find the key for the extension in HKEY_CLASSES_ROOT
            var ext = Registry.ClassesRoot.OpenSubKey(Name, false);
            if (ext == null)
                return;

            string progid = (string)ext.GetValue(null);

            RegistryKey target;
            if (progid.Length > 0)
                target = Registry.ClassesRoot.OpenSubKey(progid, true);
            else
                target = Registry.ClassesRoot.OpenSubKey(Name, true);

            if (target == null)
                return;

            // Look for PreviewDetails
            string pd = (string) target.GetValue(PreviewDetailsValueName);
            if (pd != null)
            {
                foreach (Profile p in State.Profiles)
                {
                    if (p.PreviewDetailsString == pd)
                        State.SelectedProfile = p;
                }
            }
        }

        public void RemoveHandlerFromExtension()
        {
            if (!OurHandler)
                return;

            // First, remove the handler extension key
            // Watch out for 32/64 bit issues here, as the 32-bit and 64-bit values of these are separate and isolated on 64-bit Windows,
            // the 32-bit value being under SOFTWARE\Wow6432Node.  Thus a 64-bit manager is needed to set up a 64-bit handler
            var handlers = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\PropertySystem\PropertyHandlers", true);
            if (handlers != null)
                handlers.DeleteSubKey(Name);
            this.RecordPropertyHandler (null, null);

            // Now find the key for the extension in HKEY_CLASSES_ROOT
            RegistryKey target;
            var ext = Registry.ClassesRoot.OpenSubKey(Name, false);

            // Tolerate the case where the extension has been removed since we set up a handler for it
            if (ext != null)
            {
                string progid = (string)ext.GetValue(null);

                if (progid.Length > 0)
                    target = Registry.ClassesRoot.OpenSubKey(progid, true);
                else
                    target = Registry.ClassesRoot.OpenSubKey(Name, true);

                RemoveProgidRegistryEntries(target);
            }

            // Now go after the settings in SystemFileAssociations
            // We may have created the key for the extension, but we can't tell, so we just remove the values
            var assoc = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Classes\SystemFileAssociations", true);
            target = assoc.OpenSubKey(Name, true);

            RemoveProgidRegistryEntries(target);
        }

        private void RemoveProgidRegistryEntries(RegistryKey target)
        {
            // tolerate missing entries
            if (target == null)
                return;

            target.DeleteValue(FullDetailsValueName, false);
            target.DeleteValue(PreviewDetailsValueName, false);

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
