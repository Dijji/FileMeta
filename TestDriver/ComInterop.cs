// Copyright (c) 2013, Dijji, and released under Ms-PL.  This, with other relevant licenses, can be found in the root of this distribution.
using System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
//using Microsoft.WindowsAPICodePack.Shell;
//using Microsoft.WindowsAPICodePack.Shell.PropertySystem;
using TestDriverCodePack;

namespace TestDriver
{
    //PropertyHandler
    [ComImport]
#if x64
    [Guid("D06391EE-2FEB-419B-9667-AD160D0849F3")]
#else
    [Guid("60211757-EF87-465e-B6C1-B37CF98295F9")]
#endif
    class CPropertyHandlerClass
    {
    }

    [ComImport]
    [CoClass(typeof(CPropertyHandlerClass))]
    [Guid("B7D14566-0509-4CCE-A71F-0A554233BD9B")]
    interface CPropertyHandler : IInitializeWithFile, IPropertyStore
    {
    }

    [ComImport]
    [Guid("B7D14566-0509-4CCE-A71F-0A554233BD9B")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IInitializeWithFile
    {
        void Initialize(string pszFilePath, uint grfMode);
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

    // ContextMenuHandler
    [ComImport]
#if x64
    [Guid("28D14D00-2D80-4956-9657-9D50C8BB47A5")]
#else
    [Guid("DA38301B-BE91-4397-B2C8-E27A0BD80CC5")]
#endif
    class CContextMenuHandlerClass
    {
    }

    [ComImport]
    [CoClass(typeof(CContextMenuHandlerClass))]
    [Guid("000214e8-0000-0000-c000-000000000046")]
    interface CContextMenuHandler : IShellExtInit, IContextMenu
    {
    }


    [ComImport(),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown),
     GuidAttribute("000214e8-0000-0000-c000-000000000046")]
    public interface IShellExtInit
    {
        int Initialize(IntPtr pidlFolder,
                System.Runtime.InteropServices.ComTypes.IDataObject lpdobj,
                uint /*HKEY*/ hKeyProgID);
        //[PreserveSig()]
        //int Initialize(IntPtr pidlFolder,
        //        IntPtr lpdobj,
        //        uint /*HKEY*/ hKeyProgID);
    }

    [ComImport(),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown),
     GuidAttribute("000214e4-0000-0000-c000-000000000046")]
    public interface IContextMenu
    {
        // IContextMenu methods
        [PreserveSig()]
        int QueryContextMenu(uint hmenu,
                   uint iMenu,
                   int idCmdFirst,
                   int idCmdLast,
                   uint uFlags);

        [PreserveSig()]
        void InvokeCommand(CMINVOKECOMMANDINFOEX pici);

        [PreserveSig()]
        void GetCommandString(int idcmd,
                   uint uflags,
                   int reserved,
                   string commandstring,
                   int cch);
    }

    enum ContextMenuVerbs
    {
        Export = 0,
        Import = 1,
        Delete = 2,
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct CMINVOKECOMMANDINFOEX
    {
        public int cbSize;        // Marshal.SizeOf(CMINVOKECOMMANDINFO)
        public int fMask;         // any combination of CMIC_MASK_*
        public IntPtr hwnd;          // might be NULL (indicating no owner window)
        public IntPtr lpVerb;        // either a string or MAKEINTRESOURCE(idOffset)
        public string lpParameters;      // might be NULL (indicating no parameter)
        public string lpDirectory;       // might be NULL (indicating no specific directory)
        public int nShow;         // one of SW_ values for ShowWindow() API
        public int dwHotKey;        // Optional hot key to assign to any application activated by the command. If the fMask member does not specify CMIC_MASK_HOTKEY, this member is ignored. 
        public IntPtr hIcon;            // Icon to use for any application activated by the command. If the fMask member does not specify CMIC_MASK_ICON, this member is ignored. 

        public string lpTitle;        // ASCII title.

        [MarshalAs(UnmanagedType.LPWStr)]
        public string lpVerbW;                // Unicode verb, for those commands that can use it.

        [MarshalAs(UnmanagedType.LPWStr)]
        public string lpParametersW;        // Unicode parameters, for those commands that can use it. 

        [MarshalAs(UnmanagedType.LPWStr)]
        public string lpDirectoryW;        // Unicode directory, for those commands that can use it.

        [MarshalAs(UnmanagedType.LPWStr)]
        public string lpTitleW;            // Unicode title.

        public POINT ptInvoke;                // Point where the command is invoked. This member is not valid prior to Microsoft Internet Explorer 4.0.
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct POINT
    {
        public int X;
        public int Y;

        public POINT(int x, int y)
        {
            this.X = x;
            this.Y = y;
        }

        public static implicit operator System.Drawing.Point(POINT p)
        {
            return new System.Drawing.Point(p.X, p.Y);
        }

        public static implicit operator POINT(System.Drawing.Point p)
        {
            return new POINT(p.X, p.Y);
        }
    }
}
