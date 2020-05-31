//Copied from Windows API Code Pack 1.1 and modified, mainly for public access to PropertyVariant
//Copyright (c) Microsoft Corporation.  All rights reserved.
using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace TestDriverCodePack
{
    internal enum RelativeDescriptionType
    {
        General,
        Date,
        Size,
        Count,
        Revision,
        Length,
        Duration,
        Speed,
        Rate,
        Rating,
        Priority
    }
    internal static class ShellIIDGuid
    {
        internal const string IPropertyDescription = "6F79D558-3E96-4549-A1D1-7D75D2288814";
        internal const string IPropertyDescription2 = "57D2EDED-5062-400E-B107-5DAE79FE57A6";
        internal const string IPropertyDescriptionList = "1F9FC1D0-C39B-4B26-817F-011967D3440E";
        internal const string IPropertyEnumType = "11E1FBF9-2D56-4A6B-8DB3-7CD193A471F2";
        internal const string IPropertyEnumType2 = "9B6E051C-5DDD-4321-9070-FE2ACB55E794";
        internal const string IPropertyEnumTypeList = "A99400F4-3D84-4557-94BA-1242FB2CC9A6";
    }
    [ComImport,
    Guid(ShellIIDGuid.IPropertyDescriptionList),
    InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IPropertyDescriptionList
    {
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        void GetCount(out uint pcElem);
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        void GetAt([In] uint iElem, [In] ref Guid riid, [MarshalAs(UnmanagedType.Interface)] out IPropertyDescription ppv);
    }


    [ComImport,
    Guid(ShellIIDGuid.IPropertyDescription),
    InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IPropertyDescription
    {
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        void GetPropertyKey(out PropertyKey pkey);
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        void GetCanonicalName([MarshalAs(UnmanagedType.LPWStr)] out string ppszName);
        [PreserveSig]
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        HResult GetPropertyType(out VarEnum pvartype);
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime),
        PreserveSig]
        HResult GetDisplayName(out IntPtr ppszName);
        [PreserveSig]
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        HResult GetEditInvitation(out IntPtr ppszInvite);
        [PreserveSig]
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        HResult GetTypeFlags([In] PropertyTypeOptions mask, out PropertyTypeOptions ppdtFlags);
        [PreserveSig]
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        HResult GetViewFlags(out PropertyViewOptions ppdvFlags);
        [PreserveSig]
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        HResult GetDefaultColumnWidth(out uint pcxChars);
        [PreserveSig]
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        HResult GetDisplayType(out PropertyDisplayType pdisplaytype);
        [PreserveSig]
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        HResult GetColumnState(out PropertyColumnStateOptions pcsFlags);
        [PreserveSig]
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        HResult GetGroupingRange(out PropertyGroupingRange pgr);
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        void GetRelativeDescriptionType(out RelativeDescriptionType prdt);
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        void GetRelativeDescription([In] PropVariant propvar1, [In] PropVariant propvar2, [MarshalAs(UnmanagedType.LPWStr)] out string ppszDesc1, [MarshalAs(UnmanagedType.LPWStr)] out string ppszDesc2);
        [PreserveSig]
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        HResult GetSortDescription(out PropertySortDescription psd);
        [PreserveSig]
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        HResult GetSortDescriptionLabel([In] bool fDescending, out IntPtr ppszDescription);
        [PreserveSig]
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        HResult GetAggregationType(out PropertyAggregationType paggtype);
        [PreserveSig]
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        HResult GetConditionType(out PropertyConditionType pcontype, out PropertyConditionOperation popDefault);
        [PreserveSig]
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        HResult GetEnumTypeList([In] ref Guid riid, [Out, MarshalAs(UnmanagedType.Interface)] out IPropertyEnumTypeList ppv);
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        void CoerceToCanonicalValue([In, Out] PropVariant propvar);
        [PreserveSig]
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)] // Note: this method signature may be wrong, but it is not used.
        HResult FormatForDisplay([In] PropVariant propvar, [In] ref PropertyDescriptionFormatOptions pdfFlags, [MarshalAs(UnmanagedType.LPWStr)] out string ppszDisplay);
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        HResult IsValueCanonical([In] PropVariant propvar);
    }

    [ComImport,
     Guid(ShellIIDGuid.IPropertyEnumType),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IPropertyEnumType
    {
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        void GetEnumType([Out] out PropEnumType penumtype);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        void GetValue([Out] PropVariant ppropvar);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        void GetRangeMinValue([Out] PropVariant ppropvar);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        void GetRangeSetValue([Out] PropVariant ppropvar);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        void GetDisplayText([Out, MarshalAs(UnmanagedType.LPWStr)] out string ppszDisplay);
    }

    [ComImport,
    Guid(ShellIIDGuid.IPropertyEnumTypeList),
    InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IPropertyEnumTypeList
    {
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        void GetCount([Out] out uint pctypes);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        void GetAt(
        [In] uint itype,
        [In] ref Guid riid,   // riid may be IID_IPropertyEnumType
        [Out, MarshalAs(UnmanagedType.Interface)] out IPropertyEnumType ppv);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        void GetConditionAt(
        [In] uint index,
        [In] ref Guid riid,
        out IntPtr ppv);

        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        void FindMatchingIndex(
        [In] PropVariant propvarCmp,
        [Out] out uint pnIndex);
    }

    internal static class PropertySystemNativeMethods
    {


        internal enum PropDescEnumFilter
        {
            PDEF_ALL = 0,
            PDEF_SYSTEM = 1,
            PDEF_NONSYSTEM = 2,
            PDEF_VIEWABLE = 3,
            PDEF_QUERYABLE = 4,
            PDEF_INFULLTEXTQUERY = 5,
            PDEF_COLUMN = 6
        }


        /*
                PSSTDAPI PSEnumeratePropertyDescriptions(
          _In_   PROPDESC_ENUMFILTER filterOn,
          _In_   REFIID riid,
          _Out_  void **ppv
        );*/

        [DllImport("propsys.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        internal static extern int PSEnumeratePropertyDescriptions(
            [In] PropDescEnumFilter filter,
            [In] ref Guid riid,
            out IPropertyDescriptionList ppv
        );

        [DllImport("propsys.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        internal static extern HResult PSGetPropertyDescription(
            ref PropertyKey propkey,
            ref Guid riid,
            [Out, MarshalAs(UnmanagedType.Interface)] out IPropertyDescription ppv
        );

        [DllImport("propsys.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        internal static extern int PSGetPropertyKeyFromName(
            [In, MarshalAs(UnmanagedType.LPWStr)] string pszCanonicalName,
            out PropertyKey propkey
        );

        [DllImport("propsys.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        internal static extern int PSGetNameFromPropertyKey(
            ref PropertyKey propkey,
            [Out, MarshalAs(UnmanagedType.LPWStr)] out string ppszCanonicalName
        );

        /*
        internal enum RelativeDescriptionType
        {
            General,
            Date,
            Size,
            Count,
            Revision,
            Length,
            Duration,
            Speed,
            Rate,
            Rating,
            Priority
        }





        [DllImport("propsys.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        internal static extern int PSGetPropertyDescriptionListFromString(
            [In, MarshalAs(UnmanagedType.LPWStr)] string pszPropList,
            [In] ref Guid riid,
            out IPropertyDescriptionList ppv
        );
        internal enum PROPDESC_RELATIVEDESCRIPTION_TYPE
        {
            PDRDT_GENERAL,
            PDRDT_DATE,
            PDRDT_SIZE,
            PDRDT_COUNT,
            PDRDT_REVISION,
            PDRDT_LENGTH,
            PDRDT_DURATION,
            PDRDT_SPEED,
            PDRDT_RATE,
            PDRDT_RATING,
            PDRDT_PRIORITY
        }

        [DllImport("propsys.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        internal static extern int PSGetNameFromPropertyKey(
            ref PropertyKey propkey,
            [Out, MarshalAs(UnmanagedType.LPWStr)] out string ppszCanonicalName
        );

        [DllImport("propsys.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        internal static extern HRESULT PSGetPropertyDescription(
            ref PropertyKey propkey,
            ref Guid riid,
            [Out, MarshalAs(UnmanagedType.Interface)] out IPropertyDescription ppv
        );

        [DllImport("propsys.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        internal static extern int PSGetPropertyKeyFromName(
            [In, MarshalAs(UnmanagedType.LPWStr)] string pszCanonicalName,
            out PropertyKey propkey
        );
         * */
    }
}
