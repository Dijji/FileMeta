using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;

namespace FileMetadataAssociationManager
{
    static class PropertyHandlerTest
    {
        private static Guid IPropertyStoreGuid = new Guid("886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99");

        public static bool WindowsIgnoresOurPropertyHandler(string extension)
        {
            // Create a temporary file with the right extension
            string fileFullName = Path.ChangeExtension(
                Path.GetTempPath() + Guid.NewGuid().ToString(), extension);
            using (var fs = File.Create(fileFullName)) { }

            // Get the property store that Explorer would use
            IPropertyStore ps;
            bool result = true;
            HResult hr = (HResult)SHGetPropertyStoreFromParsingName(fileFullName, IntPtr.Zero,
                GETPROPERTYSTOREFLAGS.GPS_NO_OPLOCK | GETPROPERTYSTOREFLAGS.GPS_HANDLERPROPERTIESONLY, ref IPropertyStoreGuid, out ps);
            if (hr == HResult.Ok)
            {
                // Look for the signature property value that marks the handler as ours
                PropertyKey key;
                PropVariant pv = new PropVariant();
                PropertySystemNativeMethods.PSGetPropertyKeyFromName("System.Software.ProductName", out key);
                hr = ps.GetValue(key, pv);
                if (hr == HResult.Ok && !pv.IsNullOrEmpty)
                {
                    if ((pv.Value as string) == "FileMetadata")
                        result = false;
                }
                Marshal.ReleaseComObject(ps);  // This release frees up the file for deletion
            }

            File.Delete(fileFullName);

            return result;
        }

        [DllImport("shell32.dll", SetLastError = true)]
        public static extern int SHGetPropertyStoreFromParsingName(
                [In][MarshalAs(UnmanagedType.LPWStr)] string pszPath,
                IntPtr zeroWorks,
                GETPROPERTYSTOREFLAGS flags,
                ref Guid iIdPropStore,
                [Out] out IPropertyStore propertyStore);
    }

    public enum GETPROPERTYSTOREFLAGS
    {
        // If no flags are specified (GPS_DEFAULT), a read-only property store is returned that includes properties for the file or item.
        // In the case that the shell item is a file, the property store contains:
        //     1. properties about the file from the file system
        //     2. properties from the file itself provided by the file's property handler, unless that file is offline,
        //     see GPS_OPENSLOWITEM
        //     3. if requested by the file's property handler and supported by the file system, properties stored in the
        //     alternate property store.
        //
        // Non-file shell items should return a similar read-only store
        //
        // Specifying other GPS_ flags modifies the store that is returned
        GPS_DEFAULT = 0x00000000,
        GPS_HANDLERPROPERTIESONLY = 0x00000001,   // only include properties directly from the file's property handler
        GPS_READWRITE = 0x00000002,   // Writable stores will only include handler properties
        GPS_TEMPORARY = 0x00000004,   // A read/write store that only holds properties for the lifetime of the IShellItem object
        GPS_FASTPROPERTIESONLY = 0x00000008,   // do not include any properties from the file's property handler (because the file's property handler will hit the disk)
        GPS_OPENSLOWITEM = 0x00000010,   // include properties from a file's property handler, even if it means retrieving the file from offline storage.
        GPS_DELAYCREATION = 0x00000020,   // delay the creation of the file's property handler until those properties are read, written, or enumerated
        GPS_BESTEFFORT = 0x00000040,   // For readonly stores, succeed and return all available properties, even if one or more sources of properties fails. Not valid with GPS_READWRITE.
        GPS_NO_OPLOCK = 0x00000080,   // some data sources protect the read property store with an oplock, this disables that
        GPS_MASK_VALID = 0x000000FF,
    }
    
    /// <summary>
    /// A property store
    /// </summary>
    [ComImport]
    [Guid("886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IPropertyStore
    {
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        HResult GetCount([Out] out uint propertyCount);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        HResult GetAt([In] uint propertyIndex, out PropertyKey key);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        HResult GetValue([In] ref PropertyKey key, [Out] PropVariant pv);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), PreserveSig]
        HResult SetValue([In] ref PropertyKey key, [In] PropVariant pv);

        [PreserveSig]
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        HResult Commit();
    }

    [ComImport]
    [Guid("C8E2D566-186E-4D49-BF41-6909EAD56ACC")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IPropertyStoreCapabilities
    {
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        HResult IsPropertyWritable([In]ref PropertyKey propertyKey);
    }
}
