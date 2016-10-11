﻿// Copyright (c) 2016, Dijji, and released under Ms-PL.  This, with other relevant licenses, can be found in the root of this distribution.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TestDriverAssoc
{
    class Const
    {
        public const string TestExt = ".FMfoo";

        const string OurPropertyHandlerGuid64 = "{D06391EE-2FEB-419B-9667-AD160D0849F3}";
        const string OurPropertyHandlerGuid32 = "{60211757-EF87-465e-B6C1-B37CF98295F9}";
        const string OurContextHandlerGuid64 = "{28D14D00-2D80-4956-9657-9D50C8BB47A5}";
        const string OurContextHandlerGuid32 = "{DA38301B-BE91-4397-B2C8-E27A0BD80CC5}";
        const string OtherPropertyHandlerGuid = "{a38b883c-1682-497e-97b0-0a3a9e801682}";  // Windows image

        public const string FullDetailsValueName = "FullDetails";
        public const string PreviewDetailsValueName = "PreviewDetails";
        public const string InfoTipValueName = "InfoTip";
        public const string OldFullDetailsValueName = "FileMetaOldFullDetails";
        public const string OldPreviewDetailsValueName = "FileMetaOldPreviewDetails";
        public const string OldInfoTipValueName = "FileMetaOldInfoTip";
        public const string FileMetaCustomProfileValueName = "FileMetaCustomProfile";
        public const string ContextMenuHandlersKeyName = "ShellEx\\ContextMenuHandlers\\FileMetadata";
        public const string ChainedValueName = "Chained";

#if x64
        public static string OurPropertyHandlerGuid { get { return OurPropertyHandlerGuid64; } }
        public static string OurContextHandlerGuid { get { return OurContextHandlerGuid64; } }
#elif x86
        public static string OurPropertyHandlerGuid { get { return OurPropertyHandlerGuid32; } }
        public static string OurContextHandlerGuid { get { return OurContextHandlerGuid32; } }
#endif
        const string FullDetailsOfficeProfile = "prop:System.PropGroup.Description;System.Title;System.Subject;System.Keywords;System.Category;System.Comment;System.Rating;System.PropGroup.Origin;System.Author;System.Document.LastAuthor;System.Document.RevisionNumber;System.Document.Version;System.ApplicationName;System.Company;System.Document.Manager;System.Document.DateCreated;System.Document.DateSaved;System.Document.DatePrinted;System.Document.TotalEditingTime;System.PropGroup.Content;System.ContentStatus;System.ContentType;System.Document.PageCount;System.Document.WordCount;System.Document.CharacterCount;System.Document.LineCount;System.Document.ParagraphCount;System.Document.Template;System.Document.Scale;System.Document.LinksDirty;System.Language;System.PropGroup.FileSystem;System.ItemNameDisplay;System.ItemType;System.ItemFolderPathDisplay;System.DateCreated;System.DateModified;System.Size;System.FileAttributes;System.OfflineAvailability;System.OfflineStatus;System.SharedWith;System.FileOwner;System.ComputerName";
        const string PreviewDetailsOfficeProfile = "prop:*System.DateModified;System.Author;System.Keywords;System.Rating;*System.Size;System.Title;System.Comment;System.Category;*System.Document.PageCount;System.ContentStatus;System.ContentType;*System.OfflineAvailability;*System.OfflineStatus;System.Subject;*System.DateCreated;*System.SharedWith";
        const string InfoTipOfficeProfile = "prop:System.ItemTypeText;System.Size;System.DateModified;System.Document.PageCount";

        const string FullDetailsSimpleProfile = "prop:System.PropGroup.Description;System.Title;System.Subject;System.Keywords;System.Category;System.Comment;System.Rating;System.PropGroup.Origin;System.Author;System.Document.RevisionNumber";
        const string PreviewDetailsSimpleProfile = "prop:System.Title;System.Subject;System.Keywords;System.Category;System.Comment;System.Rating;System.Author;System.Document.RevisionNumber";
        const string InfoTipSimpleProfile = "prop:System.ItemTypeText;System.Size;System.DateModified;System.Comment";

        //const string FullDetailsCustomProfile = "prop:System.PropGroup.Description;System.Keywords;System.Category;System.Comment;System.Rating;System.PropGroup.Origin;System.Author";
        //const string PreviewDetailsCustomProfile = "prop:System.Keywords;System.Category;System.Comment;System.Rating;System.Author";
        //const string InfoTipCustomProfile = "prop:System.ItemTypeText;System.Comment";

        // The following entries have to agree with the local SavedState.xml
        const string FullDetailsCustomProfileTest = "prop:System.PropGroup.Description;System.Title;System.Subject;System.Keywords;System.Category;System.Comment;System.Rating;System.PropGroup.Origin;System.Author;System.Document.RevisionNumber";
        const string PreviewDetailsCustomProfileTest = "prop:System.Title;System.Subject;System.Keywords;System.Category;System.Comment;System.Rating;System.Author;System.Document.RevisionNumber";
        const string InfoTipCustomProfileTest = "prop:System.Comment;System.Rating";

        const string FullDetailsCustomProfileBmp = "prop:System.PropGroup.Image;System.Image.Dimensions;System.Image.HorizontalSize;System.Image.VerticalSize;System.Image.BitDepth;System.PropGroup.FileSystem;System.ItemNameDisplay;System.ItemType;System.ItemFolderPathDisplay;System.DateCreated;System.DateModified;System.Size;System.FileAttributes;System.OfflineAvailability;System.OfflineStatus;System.SharedWith;System.FileOwner;System.ComputerName;System.PropGroup.Description;System.Title;System.Subject;System.Keywords;System.Category;System.Comment;System.Rating;System.PropGroup.Origin;System.Author;System.Document.RevisionNumber";
        const string PreviewDetailsCustomProfileBmp = "prop:*System.DateModified;*System.Image.Dimensions;*System.Size;*System.OfflineAvailability;*System.OfflineStatus;*System.DateCreated;*System.SharedWith;System.Title;System.Subject;System.Keywords;System.Category;System.Comment;System.Rating;System.Author;System.Document.RevisionNumber";
        const string InfoTipCustomProfileBmp = "prop:System.ItemType;*System.DateModified;*System.Image.Dimensions;*System.Size;System.ItemTypeText;System.Comment";

        public static RegState V13BuiltIn = new RegState
        {
            ClsidFullDetails = FullDetailsSimpleProfile,
            ClsidPreviewDetails = PreviewDetailsSimpleProfile,
            ClsidContextMenuHandler = OurContextHandlerGuid,
            SystemFullDetails = FullDetailsSimpleProfile,
            SystemPreviewDetails = PreviewDetailsSimpleProfile,
            SystemContextMenuHandler = OurContextHandlerGuid,
            PropertyHandler = OurPropertyHandlerGuid,
        };
        public static RegState V13CustomTest = new RegState
        {
            ClsidFullDetails = FullDetailsCustomProfileTest,
            ClsidPreviewDetails = PreviewDetailsCustomProfileTest,
            ClsidCustomProfile = "test",
            SystemFullDetails = FullDetailsCustomProfileTest,
            SystemPreviewDetails = PreviewDetailsCustomProfileTest,
            SystemCustomProfile = "test",
            PropertyHandler = OurPropertyHandlerGuid,
        };
        public static RegState V14BuiltIn = new RegState
        {
            SystemFullDetails = FullDetailsSimpleProfile,
            SystemPreviewDetails = PreviewDetailsSimpleProfile,
            SystemInfoTip = InfoTipSimpleProfile,
            SystemContextMenuHandler = OurContextHandlerGuid,
            PropertyHandler = OurPropertyHandlerGuid,
        };
        public static RegState V14CustomTest = new RegState
        {
            SystemFullDetails = FullDetailsCustomProfileTest,
            SystemPreviewDetails = PreviewDetailsCustomProfileTest,
            SystemInfoTip = InfoTipCustomProfileTest,
            SystemCustomProfile = "test",
            SystemContextMenuHandler = OurContextHandlerGuid,
            PropertyHandler = OurPropertyHandlerGuid,
        };
        public static RegState V14ExtendedBmp = new RegState
        {
            SystemFullDetails = FullDetailsCustomProfileBmp,
            SystemPreviewDetails = PreviewDetailsCustomProfileBmp,
            SystemInfoTip = InfoTipCustomProfileBmp,
            SystemOldFullDetails = FullDetailsSimpleProfile,
            SystemOldPreviewDetails = PreviewDetailsSimpleProfile,
            SystemOldInfoTip = InfoTipSimpleProfile,
            SystemCustomProfile = ".bmp",
            SystemContextMenuHandler = OurContextHandlerGuid,
            PropertyHandler = OurPropertyHandlerGuid,
            ChainedPropertyHandler = OtherPropertyHandlerGuid,
        };
        // What should be left after removing V14Extended
        public static RegState V14UnExtended = new RegState
        {
            SystemFullDetails = FullDetailsSimpleProfile,
            SystemPreviewDetails = PreviewDetailsSimpleProfile,
            SystemInfoTip = InfoTipSimpleProfile,
            PropertyHandler = OtherPropertyHandlerGuid,
        };
        public static RegState V15BuiltIn = new RegState
        {
            SystemFullDetails = FullDetailsSimpleProfile,
            SystemPreviewDetails = PreviewDetailsSimpleProfile,
            SystemInfoTip = InfoTipSimpleProfile,
            SystemContextMenuHandler = OurContextHandlerGuid,
            PropertyHandler = OurPropertyHandlerGuid,
#if x64
            PropertyHandler32 = OurPropertyHandlerGuid32,
#endif
        };
        public static RegState V15CustomTest = new RegState
        {
            SystemFullDetails = FullDetailsCustomProfileTest,
            SystemPreviewDetails = PreviewDetailsCustomProfileTest,
            SystemInfoTip = InfoTipCustomProfileTest,
            SystemCustomProfile = "test",
            SystemContextMenuHandler = OurContextHandlerGuid,
            PropertyHandler = OurPropertyHandlerGuid,
#if x64
            PropertyHandler32 = OurPropertyHandlerGuid32,
#endif
        };
#if x64
        // This is an initial state where there is no 64-bit property handler installed, but there is a 32-bit property handler
        public static RegState V15InitialOther32 = new RegState
        {
             PropertyHandler32 = OtherPropertyHandlerGuid,
        };
#endif
        // This is what we should find after we have added our handler in the above case
        public static RegState V15CustomTestOther32 = new RegState
        {
            SystemFullDetails = FullDetailsCustomProfileTest,
            SystemPreviewDetails = PreviewDetailsCustomProfileTest,
            SystemInfoTip = InfoTipCustomProfileTest,
            SystemCustomProfile = "test",
            SystemContextMenuHandler = OurContextHandlerGuid,
            PropertyHandler = OurPropertyHandlerGuid,
#if x64
            PropertyHandler32 = OtherPropertyHandlerGuid,
#endif
        };
        public static RegState V15ExtendedBmp = new RegState
        {
            SystemFullDetails = FullDetailsCustomProfileBmp,
            SystemPreviewDetails = PreviewDetailsCustomProfileBmp,
            SystemInfoTip = InfoTipCustomProfileBmp,
            SystemOldFullDetails = FullDetailsSimpleProfile,
            SystemOldPreviewDetails = PreviewDetailsSimpleProfile,
            SystemOldInfoTip = InfoTipSimpleProfile,
            SystemCustomProfile = ".bmp",
            SystemContextMenuHandler = OurContextHandlerGuid,
            PropertyHandler = OurPropertyHandlerGuid,
            ChainedPropertyHandler = OtherPropertyHandlerGuid,
#if x64
            PropertyHandler32 = OurPropertyHandlerGuid32,
            ChainedPropertyHandler32 = OtherPropertyHandlerGuid,
#endif
        };
        // What should be left after removing V15Extended
        public static RegState V15UnExtended = new RegState
        {
            SystemFullDetails = FullDetailsSimpleProfile,
            SystemPreviewDetails = PreviewDetailsSimpleProfile,
            SystemInfoTip = InfoTipSimpleProfile,
            PropertyHandler = OtherPropertyHandlerGuid,
#if x64
            PropertyHandler32 = OtherPropertyHandlerGuid,
#endif
        };
        public static RegState V15ExtendedBmpClsid = new RegState
        {
            SystemFullDetails = FullDetailsCustomProfileBmp,
            SystemPreviewDetails = PreviewDetailsCustomProfileBmp,
            SystemInfoTip = InfoTipCustomProfileBmp,
            ClsidOldFullDetails = FullDetailsSimpleProfile,
            ClsidOldPreviewDetails = PreviewDetailsSimpleProfile,
            ClsidOldInfoTip = InfoTipSimpleProfile,
            SystemCustomProfile = ".bmp",
            SystemContextMenuHandler = OurContextHandlerGuid,
            PropertyHandler = OurPropertyHandlerGuid,
            ChainedPropertyHandler = OtherPropertyHandlerGuid,
#if x64
            PropertyHandler32 = OurPropertyHandlerGuid32,
            ChainedPropertyHandler32 = OtherPropertyHandlerGuid,
#endif
        };
        // What should be left after removing V15Extended
        public static RegState V15UnExtendedClsid = new RegState
        {
            ClsidFullDetails = FullDetailsSimpleProfile,
            ClsidPreviewDetails = PreviewDetailsSimpleProfile,
            ClsidInfoTip = InfoTipSimpleProfile,
            PropertyHandler = OtherPropertyHandlerGuid,
#if x64
            PropertyHandler32 = OtherPropertyHandlerGuid,
#endif
        };
        public static RegState V15ExtendedBmpBoth = new RegState
        {
            SystemFullDetails = FullDetailsCustomProfileBmp,
            SystemPreviewDetails = PreviewDetailsCustomProfileBmp,
            SystemInfoTip = InfoTipCustomProfileBmp,
            SystemOldFullDetails = FullDetailsSimpleProfile,
            SystemOldPreviewDetails = PreviewDetailsSimpleProfile,
            SystemOldInfoTip = InfoTipSimpleProfile,
            ClsidOldFullDetails = FullDetailsSimpleProfile,
            ClsidOldPreviewDetails = PreviewDetailsSimpleProfile,
            ClsidOldInfoTip = InfoTipSimpleProfile,
            SystemCustomProfile = ".bmp",
            SystemContextMenuHandler = OurContextHandlerGuid,
            PropertyHandler = OurPropertyHandlerGuid,
            ChainedPropertyHandler = OtherPropertyHandlerGuid,
#if x64
            PropertyHandler32 = OurPropertyHandlerGuid32,
            ChainedPropertyHandler32 = OtherPropertyHandlerGuid,
#endif
        };
        // What should be left after removing V15Extended
        public static RegState V15UnExtendedBoth = new RegState
        {
            ClsidFullDetails = FullDetailsSimpleProfile,
            ClsidPreviewDetails = PreviewDetailsSimpleProfile,
            ClsidInfoTip = InfoTipSimpleProfile,
            SystemFullDetails = FullDetailsSimpleProfile,
            SystemPreviewDetails = PreviewDetailsSimpleProfile,
            SystemInfoTip = InfoTipSimpleProfile,
            PropertyHandler = OtherPropertyHandlerGuid,
#if x64
            PropertyHandler32 = OtherPropertyHandlerGuid,
#endif
        }; 
        // Below is the profile being merged in in the registry states that follow
        //<Profile>
        //  <Name>tiny</Name>
        //  <FullDetailsString>prop:System.PropGroup.Description;System.Title;System.Comment;System.PropGroup.Origin;System.Author</FullDetailsString>
        //  <PreviewDetailsString>prop:System.Comment;System.Author</PreviewDetailsString>
        //  <InfoTipString>prop:System.Comment</InfoTipString>
        //</Profile>

        // Starting state
        public static RegState V15UnExtendedTiny = new RegState
        {
            SystemFullDetails = "prop:System.PropGroup.Description;System.Subject;System.Keywords;System.PropGroup.Content;System.ContentStatus",
            SystemPreviewDetails = "prop:System.Description;System.ContentStatus",
            SystemInfoTip = "prop:System.Description",
            PropertyHandler = OtherPropertyHandlerGuid,
#if x64
            PropertyHandler32 = OtherPropertyHandlerGuid,
#endif
        };
        // Extended with the tiny profile
        public static RegState V15ExtendedTiny = new RegState
        {
            SystemFullDetails = "prop:System.PropGroup.Description;System.Subject;System.Keywords;System.Title;System.Comment;System.PropGroup.Content;System.ContentStatus;System.PropGroup.Origin;System.Author",
            SystemPreviewDetails = "prop:System.Description;System.ContentStatus;System.Comment;System.Author",
            SystemInfoTip = "prop:System.Description;System.Comment",
            SystemOldFullDetails = "prop:System.PropGroup.Description;System.Subject;System.Keywords;System.PropGroup.Content;System.ContentStatus",
            SystemOldPreviewDetails = "prop:System.Description;System.ContentStatus",
            SystemOldInfoTip = "prop:System.Description",
            SystemCustomProfile = ".FMfoo",
            SystemContextMenuHandler = OurContextHandlerGuid,
            PropertyHandler = OurPropertyHandlerGuid,
            ChainedPropertyHandler = OtherPropertyHandlerGuid,
#if x64
            PropertyHandler32 = OurPropertyHandlerGuid32,
            ChainedPropertyHandler32 = OtherPropertyHandlerGuid,
#endif
        };
    }
}
