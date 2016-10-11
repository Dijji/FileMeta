// Copyright (c) 2016, Dijji, and released under Ms-PL.  This, with other relevant licenses, can be found in the root of this distribution.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Win32;

namespace TestDriverAssoc
{
    // Everything that File Meta puts into the registry for an extension,
    // characterising a registry state
    struct RegState
    {
        public string ClsidFullDetails;
        public string ClsidPreviewDetails;
        public string ClsidInfoTip;
        public string ClsidOldFullDetails;
        public string ClsidOldPreviewDetails;
        public string ClsidOldInfoTip;
        public string ClsidCustomProfile;
        public string ClsidContextMenuHandler;
        public string SystemFullDetails;
        public string SystemPreviewDetails;
        public string SystemInfoTip;
        public string SystemOldFullDetails;
        public string SystemOldPreviewDetails;
        public string SystemOldInfoTip;
        public string SystemCustomProfile;
        public string SystemContextMenuHandler;
        public string PropertyHandler;
        public string ChainedPropertyHandler;
#if x64
        public string PropertyHandler32;
        public string ChainedPropertyHandler32;
#endif

        public static bool operator ==(RegState r1, RegState r2)
        {
            return     
                r1.ClsidFullDetails == r2.ClsidFullDetails &&
                r1.ClsidPreviewDetails == r2.ClsidPreviewDetails &&
                r1.ClsidInfoTip == r2.ClsidInfoTip &&
                r1.ClsidOldFullDetails == r2.ClsidOldFullDetails &&
                r1.ClsidOldPreviewDetails == r2.ClsidOldPreviewDetails &&
                r1.ClsidOldInfoTip == r2.ClsidOldInfoTip &&
                r1.ClsidCustomProfile == r2.ClsidCustomProfile &&
                r1.ClsidContextMenuHandler == r2.ClsidContextMenuHandler &&
                r1.SystemFullDetails == r2.SystemFullDetails &&
                r1.SystemPreviewDetails == r2.SystemPreviewDetails &&
                r1.SystemInfoTip == r2.SystemInfoTip &&
                r1.SystemOldFullDetails == r2.SystemOldFullDetails &&
                r1.SystemOldPreviewDetails == r2.SystemOldPreviewDetails &&
                r1.SystemOldInfoTip == r2.SystemOldInfoTip &&
                r1.SystemCustomProfile == r2.SystemCustomProfile &&
                r1.SystemContextMenuHandler == r2.SystemContextMenuHandler &&
        #if x64
                r1.PropertyHandler32 == r2.PropertyHandler32 &&
                r1.ChainedPropertyHandler32 == r2.ChainedPropertyHandler32 &&
        #endif 
                r1.PropertyHandler == r2.PropertyHandler &&
                r1.PropertyHandler == r2.PropertyHandler;
        }

        public static bool operator !=(RegState r1, RegState r2)
        {
            return !(r1 == r2);
        }

        // Populate the state by reading the registry for the given extension
        public void Read(string ext)
        {
            using (RegistryKey target = GetHKCRProfileKey(ext))
            {
                if (target != null)
                {
                    ClsidFullDetails = target.GetValue(Const.FullDetailsValueName) as string;
                    ClsidPreviewDetails = target.GetValue(Const.PreviewDetailsValueName) as string;
                    ClsidInfoTip = target.GetValue(Const.InfoTipValueName) as string;
                    ClsidOldFullDetails = target.GetValue(Const.OldFullDetailsValueName) as string;
                    ClsidOldPreviewDetails = target.GetValue(Const.OldPreviewDetailsValueName) as string;
                    ClsidOldInfoTip = target.GetValue(Const.OldInfoTipValueName) as string;
                    ClsidCustomProfile = target.GetValue(Const.FileMetaCustomProfileValueName) as string;

                    using (RegistryKey keyHandler = target.OpenSubKey(Const.ContextMenuHandlersKeyName, false))
                    {
                        if (keyHandler != null)
                            ClsidContextMenuHandler = keyHandler.GetValue(null) as string;
                    }
                }
            }
            using (RegistryKey target = GetSystemFileAssociationsProfileKey(ext))
            {
                if (target != null)
                {
                    SystemFullDetails = target.GetValue(Const.FullDetailsValueName) as string;
                    SystemPreviewDetails = target.GetValue(Const.PreviewDetailsValueName) as string;
                    SystemInfoTip = target.GetValue(Const.InfoTipValueName) as string;
                    SystemOldFullDetails = target.GetValue(Const.OldFullDetailsValueName) as string;
                    SystemOldPreviewDetails = target.GetValue(Const.OldPreviewDetailsValueName) as string;
                    SystemOldInfoTip = target.GetValue(Const.OldInfoTipValueName) as string;
                    SystemCustomProfile = target.GetValue(Const.FileMetaCustomProfileValueName) as string;

                    using (RegistryKey keyHandler = target.OpenSubKey(Const.ContextMenuHandlersKeyName, false))
                    {
                        if (keyHandler != null)
                            SystemContextMenuHandler = keyHandler.GetValue(null) as string;
                    }
                }
            }
            using (RegistryKey handlers = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\PropertySystem\PropertyHandlers", false))
            {
                using (RegistryKey handler = handlers.OpenSubKey(ext, false))
                {
                    if (handler != null)
                    {
                        PropertyHandler = handler.GetValue(null) as string;
                        ChainedPropertyHandler = handler.GetValue(Const.ChainedValueName) as string;
                    }
                }
            }
#if x64
            using (RegistryKey handlers = RegistryExtensions.OpenBaseKey(RegistryHive.LocalMachine, RegistryExtensions.RegistryHiveType.X86).
                                            OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\PropertySystem\PropertyHandlers", false))
            {
                using (RegistryKey handler = handlers.OpenSubKey(ext, false))
                {
                    if (handler != null)
                    {
                        PropertyHandler32 = handler.GetValue(null) as string;
                        ChainedPropertyHandler32 = handler.GetValue(Const.ChainedValueName) as string;
                    }
                }
            }
#endif
        }

        public static void CreateMinimalExtension(string ext)
        {
            using (RegistryKey target = GetHKCRProfileKey(ext, true))
            { 
            }
        }
        
        // Zap our state into the registry for the given extension
        // Do not use this for real extensions, it is intended only for setting up test extensions
        public void Zap(string ext)
        {
            using (RegistryKey target = GetHKCRProfileKey(ext, true))
            {
                SetOrDeleteValue(target, Const.FullDetailsValueName, ClsidFullDetails);
                SetOrDeleteValue(target, Const.PreviewDetailsValueName, ClsidPreviewDetails);
                SetOrDeleteValue(target, Const.InfoTipValueName, ClsidInfoTip);
                SetOrDeleteValue(target, Const.OldFullDetailsValueName, ClsidOldFullDetails);
                SetOrDeleteValue(target, Const.OldPreviewDetailsValueName, ClsidOldPreviewDetails);
                SetOrDeleteValue(target, Const.OldInfoTipValueName, ClsidOldInfoTip);
                SetOrDeleteValue(target, Const.FileMetaCustomProfileValueName, ClsidCustomProfile);

                if (ClsidContextMenuHandler != null)
                    using (RegistryKey keyHandler = target.CreateSubKey(Const.ContextMenuHandlersKeyName))
                    {
                        keyHandler.SetValue(null, ClsidContextMenuHandler);
                    }
                else
                    target.DeleteSubKey(Const.ContextMenuHandlersKeyName, false);
            }

            using (RegistryKey target = GetSystemFileAssociationsProfileKey(ext, true))
            {
                SetOrDeleteValue(target, Const.FullDetailsValueName, SystemFullDetails);
                SetOrDeleteValue(target, Const.PreviewDetailsValueName, SystemPreviewDetails);
                SetOrDeleteValue(target, Const.InfoTipValueName, SystemInfoTip);
                SetOrDeleteValue(target, Const.OldFullDetailsValueName, SystemOldFullDetails);
                SetOrDeleteValue(target, Const.OldPreviewDetailsValueName, SystemOldPreviewDetails);
                SetOrDeleteValue(target, Const.OldInfoTipValueName, SystemOldInfoTip);
                SetOrDeleteValue(target, Const.FileMetaCustomProfileValueName, SystemCustomProfile);

                if (SystemContextMenuHandler != null)
                    using (RegistryKey keyHandler = target.CreateSubKey(Const.ContextMenuHandlersKeyName))
                    {
                        keyHandler.SetValue(null, SystemContextMenuHandler);
                    }
                else
                    target.DeleteSubKey(Const.ContextMenuHandlersKeyName, false);
            }

            using (RegistryKey handlers = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\PropertySystem\PropertyHandlers", true))
            {
                if (PropertyHandler != null || ChainedPropertyHandler != null)
                    using (RegistryKey handler = handlers.CreateSubKey(ext))
                    {
                        SetOrDeleteValue(handler, null, PropertyHandler);
                        SetOrDeleteValue(handler, Const.ChainedValueName, ChainedPropertyHandler);
                    }
                else
                    handlers.DeleteSubKey(ext, false);
            }
#if x64
            using (RegistryKey handlers = RegistryExtensions.OpenBaseKey(RegistryHive.LocalMachine, RegistryExtensions.RegistryHiveType.X86, true).
                                            OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\PropertySystem\PropertyHandlers", true))
            {
                if (PropertyHandler32 != null || ChainedPropertyHandler32 != null)
                    using (RegistryKey handler = handlers.CreateSubKey(ext))
                    {
                        SetOrDeleteValue(handler, null, PropertyHandler32);
                        SetOrDeleteValue(handler, Const.ChainedValueName, ChainedPropertyHandler32);
                    }
                else
                    handlers.DeleteSubKey(ext, false);
            }
#endif
        }

        // Wipe the entries we use from the registry for the given extension
        // Do not use this for real extensions, it is intended only for cleaning up test extensions
        public static void Wipe(string ext)
        {
            if (Registry.ClassesRoot.GetSubKeyNames().Contains(ext))
                 Registry.ClassesRoot.DeleteSubKeyTree(ext);

            using (RegistryKey assoc = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Classes\SystemFileAssociations", true))
            {
                if (assoc.GetSubKeyNames().Contains(ext))
                    assoc.DeleteSubKeyTree(ext);
            }

            using (RegistryKey handlers = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\PropertySystem\PropertyHandlers", true))
            {
                if (handlers.GetSubKeyNames().Contains(ext))
                    handlers.DeleteSubKeyTree(ext);
            }

#if x64
            using (RegistryKey handlers = RegistryExtensions.OpenBaseKey(RegistryHive.LocalMachine, RegistryExtensions.RegistryHiveType.X86, true).
                                             OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\PropertySystem\PropertyHandlers", true))

            {
                if (handlers.GetSubKeyNames().Contains(ext))
                    handlers.DeleteSubKeyTree(ext);
            } 
#endif
        }

        private void SetOrDeleteValue(RegistryKey key, string valueName, string value )
        {
            if (value == null)
                key.DeleteValue(valueName, false);
            else
                key.SetValue(valueName, value);
        }

        private static RegistryKey GetHKCRProfileKey(string ext, bool bWritable = false)
        {
            // Find the key for the extension in HKEY_CLASSES_ROOT
            string progid;

            using (RegistryKey extKey = Registry.ClassesRoot.OpenSubKey(ext, false))
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
                targetName = ext;

            if (bWritable)
                return Registry.ClassesRoot.CreateSubKey(targetName);
            else
                return Registry.ClassesRoot.OpenSubKey(targetName, false);
        }

        private RegistryKey GetSystemFileAssociationsProfileKey(string ext, bool bWritable = false)
        {
            using (RegistryKey assoc = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Classes\SystemFileAssociations", bWritable))
            {
                if (assoc == null)
                    return null;

                if (bWritable)
                    return assoc.CreateSubKey(ext);
                else
                    return assoc.OpenSubKey(ext, false);
            }
        }
    }
}
